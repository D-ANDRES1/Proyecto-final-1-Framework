using System;
using System.Data;
using System.IO;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using Microsoft.Office.Core;
using Newtonsoft.Json.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
//https://drive.google.com/drive/folders/1qxfDSkRZyIS7iLl9M3tNFrilbHv66rAG?usp=drive_link explicacion del video
//// "GroqApiKey": "",
//"GroqEndpoint": "https://api.groq.com/openai/v1/chat/completions",
//    "GroqModel": "llama-3.1-8b-instant",
//    "ConnectionStrings": {
//    "DefaultConnectio!#$n": "Server=;Database=dbInvestigacion;Integrated Security=true;TrustServerCertificate=true;"
//    },
//    "OutputPath": "C:\\Users\\HP\\Documents"
//}

namespace ProyectoIA
{
    // Valores numéricos equivalentes
    public enum PpSlideLayout
    {
        ppLayoutText = 2
    }

    public enum PpSaveAsFileType
    {
        ppSaveAsDefault = 11
    }
    public partial class Form1 : Form
    {
        private static readonly HttpClient _httpClient = new HttpClient();
        private int _ultimosTokens;
        private double _ultimoTiempo;
        private string _ultimaRespuesta;

        public Form1()
        {
            InitializeComponent();
            
        }

        private async void btnGenerar_Click(object sender, EventArgs e)
        {
            string pregunta = txtPregunta.Text;

            btnGenerar.Enabled = false;
            btnGuardar.Enabled = false;

            try
            {
                if (string.IsNullOrWhiteSpace(pregunta))
                {
                    MessageBox.Show("Ingrese una pregunta válida",
                                    "Validación",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                    return;
                }

                var outputPath = AppSettings.OutputPath;
                Directory.CreateDirectory(outputPath);

                // Obtener respuesta de la API
                var (respuesta, tokens, tiempo) = await ConsultarIA(
                    pregunta,
                    AppSettings.GroqEndpoint,
                    AppSettings.GroqApiKey,
                    AppSettings.GroqModel
                );

                // Actualizar controles en el hilo de UI
                this.Invoke((MethodInvoker)delegate
                {
                    _ultimosTokens = tokens;
                    _ultimoTiempo = tiempo;
                    _ultimaRespuesta = respuesta;

                    txtRespuesta.Text = respuesta;
                    lblTokens.Text = $"Tokens: {tokens} | Tiempo: {tiempo:F2}s";

                    // Forzar refresco visual
                    statusStrip1.Refresh();
                });

                // Generar documentos
                var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
                GenerarWord("Respuesta IA", respuesta, "Fin del documento",
                    Path.Combine(outputPath, $"respuesta_{timestamp}.docx"));

                GenerarPowerPoint("Resultado IA", respuesta,
                    Path.Combine(outputPath, $"presentacion_{timestamp}.pptx"));

                // Guardar en BD
                

                MessageBox.Show("Proceso completado exitosamente",
                                "Éxito",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}",
                                "Error",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
            }
            finally
            {
                btnGenerar.Enabled = true;
                btnGuardar.Enabled = !string.IsNullOrEmpty(_ultimaRespuesta);
            }
        }

        private async Task<(string respuesta, int tokens, double tiempo)> ConsultarIA(
            string pregunta, string endpoint, string apiKey, string modelo)
        {
            try
            {
                var payload = new
                {
                    model = modelo,
                    messages = new[] { new { role = "user", content = pregunta } }
                };

                var content = new StringContent(
                    Newtonsoft.Json.JsonConvert.SerializeObject(payload),
                    Encoding.UTF8,
                    "application/json"
                );

                _httpClient.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);

                var stopwatch = System.Diagnostics.Stopwatch.StartNew();
                var response = await _httpClient.PostAsync(endpoint, content);
                stopwatch.Stop();

                response.EnsureSuccessStatusCode();
                var json = await response.Content.ReadAsStringAsync();

                JObject obj = JObject.Parse(json);
                return (
                    texto: obj["choices"]?[0]?["message"]?["content"]?.ToString()
                        ?? throw new Exception("Respuesta inválida de la API"),
                    totalTokens: obj["usage"]?["total_tokens"]?.Value<int>() ?? 0,
                    tiempo: stopwatch.Elapsed.TotalSeconds
                );
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en consulta IA: {ex.Message}", ex);
            }
        }

        private void GenerarWord(string titulo, string contenido, string pie, string ruta)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                wordApp = new Word.Application { Visible = false };
                doc = wordApp.Documents.Add();

                Word.Paragraph titlePara = doc.Content.Paragraphs.Add();
                titlePara.Range.Text = titulo;
                titlePara.Range.Font.Bold = 1;
                titlePara.Range.Font.Size = 16;
                titlePara.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                titlePara.Range.InsertParagraphAfter();

                Word.Paragraph contentPara = doc.Content.Paragraphs.Add();
                contentPara.Range.Text = contenido;
                contentPara.Range.Font.Size = 12;
                contentPara.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                contentPara.Range.InsertParagraphAfter();

                doc.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);

                Word.Paragraph footerPara = doc.Content.Paragraphs.Add();
                footerPara.Range.Text = pie;
                footerPara.Range.Font.Italic = 1;
                footerPara.Range.Font.Size = 10;
                footerPara.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                footerPara.Range.InsertParagraphAfter();

                doc.SaveAs2(ruta);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en Word: {ex.Message}", ex);
            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();
                    Marshal.ReleaseComObject(doc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void GenerarPowerPoint(string titulo, string contenido, string ruta)
        {
            PowerPoint.Application pptApp = null;
            PowerPoint.Presentation presentation = null;
            PowerPoint.Slides slides = null;
            PowerPoint.Slide slide = null;

            try
            {
                pptApp = new PowerPoint.Application { Visible = MsoTriState.msoTrue };
                presentation = pptApp.Presentations.Add(MsoTriState.msoTrue);
                slides = presentation.Slides;

                // Crear diapositiva con diseño ppLayoutText (2 placeholders: título y contenido)
                slide = slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutText);

                // ====== PERSONALIZACIÓN DEL TÍTULO (Shape 1) ======
                PowerPoint.Shape shapeTitulo = slide.Shapes[1];
                shapeTitulo.TextFrame.TextRange.Text = titulo;

                // Estilo del título
                shapeTitulo.TextFrame.TextRange.Font.Size = 32;
                shapeTitulo.TextFrame.TextRange.Font.Name = "Arial";
                shapeTitulo.TextFrame.TextRange.Font.Color.RGB = RGB(41, 80, 150); // Azul
                shapeTitulo.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;

                // ====== PERSONALIZACIÓN DEL CONTENIDO (Shape 2) ======
                PowerPoint.Shape shapeContenido = slide.Shapes[2];
                shapeContenido.TextFrame.TextRange.Text = contenido;

                // Estilo del contenido
                shapeContenido.TextFrame.TextRange.Font.Size = 20;
                shapeContenido.TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0); // Negro

                // Viñetas automáticas
                shapeContenido.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
                shapeContenido.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 8226; // Código Unicode para bullet (•)

                // ====== FONDO DE LA DIAPOSITIVA ======
                slide.FollowMasterBackground = MsoTriState.msoFalse;
                slide.Background.Fill.ForeColor.RGB = RGB(242, 242, 242); // Gris claro

                // ====== GUARDAR ======
                presentation.SaveAs(ruta, PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoFalse);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error en PowerPoint: {ex.Message}", ex);
            }
            finally
            {
                // Liberar recursos en orden inverso
                if (slide != null) Marshal.ReleaseComObject(slide);
                if (slides != null) Marshal.ReleaseComObject(slides);
                if (presentation != null)
                {
                    presentation.Close();
                    Marshal.ReleaseComObject(presentation);
                }
                if (pptApp != null)
                {
                    pptApp.Quit();
                    Marshal.ReleaseComObject(pptApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Método auxiliar para colores RGB
        private int RGB(int rojo, int verde, int azul)
        {
            return (azul << 16) | (verde << 8) | rojo;
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_ultimaRespuesta))
                {
                    MessageBox.Show("No hay respuesta para guardar");
                    return;
                }

                GuardarEnBD(
                    AppSettings.ConnectionString,
                    txtPregunta.Text,
                    _ultimaRespuesta,
                    _ultimosTokens,
                    _ultimoTiempo
                );

                MessageBox.Show("Registro guardado exitosamente en la base de datos");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error al guardar: {ex.Message}");
            }
        }

        private void GuardarEnBD(string connStr, string pregunta, string respuesta, int tokens, double tiempo)
        {
            if (string.IsNullOrWhiteSpace(pregunta))
                throw new ArgumentException("La pregunta no puede estar vacía");

            if (string.IsNullOrWhiteSpace(respuesta))
                throw new ArgumentException("La respuesta generada es inválida");

            try
            {
                using (var conn = new SqlConnection(connStr))
                {
                    conn.Open();
                    var query = @"INSERT INTO RespuestasIA 
                          (Pregunta, Respuesta, Tokens, TiempoRespuesta, Fecha) 
                          VALUES (@p, @r, @t, @ti, GETDATE())";

                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.CommandTimeout = 30; // Timeout opcional
                        cmd.Parameters.Add("@p", SqlDbType.NVarChar, 500).Value = pregunta;
                        cmd.Parameters.Add("@r", SqlDbType.NVarChar, -1).Value = respuesta; // -1 = NVARCHAR(MAX)
                        cmd.Parameters.Add("@t", SqlDbType.Int).Value = tokens;
                        cmd.Parameters.Add("@ti", SqlDbType.Float).Value = tiempo;

                        cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (SqlException sqlEx)
            {
                throw new Exception($"Error de SQL: {sqlEx.Message}", sqlEx);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error al guardar en BD: {ex.Message}", ex);
            }
        }
    }

        public static class AppSettings
    {
        private static readonly JObject config;

        static AppSettings()
        {
            try
            {
                if (!File.Exists("appsettings.json"))
                    throw new FileNotFoundException("Archivo appsettings.json no encontrado");

                config = JObject.Parse(File.ReadAllText("appsettings.json"));

                if (string.IsNullOrEmpty(GroqApiKey))
                    throw new ArgumentException("GroqApiKey no configurada");

                if (string.IsNullOrEmpty(ConnectionString))
                    throw new ArgumentException("Cadena de conexión no configurada");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error de configuración: {ex.Message}");
                Environment.Exit(1);
            }
        }

        public static string GroqApiKey => config["GroqApiKey"]?.ToString();
        public static string GroqEndpoint => config["GroqEndpoint"]?.ToString();
        public static string GroqModel => config["GroqModel"]?.ToString();
        public static string OutputPath => config["OutputPath"]?.ToString();
        public static string ConnectionString => config["ConnectionStrings"]?["DefaultConnection"]?.ToString();
    }
}
