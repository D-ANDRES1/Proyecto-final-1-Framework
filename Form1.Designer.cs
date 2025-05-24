// Form1.Designer.cs
using System.Windows.Forms;

namespace ProyectoIA
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        private TextBox txtPregunta;
        private Button btnGenerar;
        private RichTextBox txtRespuesta;
        private Label lblRespuesta;
        private Button btnGuardar;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel lblTokens;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.txtPregunta = new System.Windows.Forms.TextBox();
            this.btnGenerar = new System.Windows.Forms.Button();
            this.txtRespuesta = new System.Windows.Forms.RichTextBox();
            this.lblRespuesta = new System.Windows.Forms.Label();
            this.btnGuardar = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.lblTokens = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtPregunta
            // 
            this.txtPregunta.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtPregunta.Location = new System.Drawing.Point(12, 12);
            this.txtPregunta.Multiline = true;
            this.txtPregunta.Name = "txtPregunta";
            this.txtPregunta.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtPregunta.Size = new System.Drawing.Size(460, 100);
            this.txtPregunta.TabIndex = 0;
            // 
            // btnGenerar
            // 
            this.btnGenerar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGenerar.Location = new System.Drawing.Point(397, 118);
            this.btnGenerar.Name = "btnGenerar";
            this.btnGenerar.Size = new System.Drawing.Size(75, 23);
            this.btnGenerar.TabIndex = 1;
            this.btnGenerar.Text = "Generar";
            this.btnGenerar.UseVisualStyleBackColor = true;
            btnGenerar.Click += btnGenerar_Click;
            // !#$
            // txtRespuesta
            // 
            this.txtRespuesta.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtRespuesta.Location = new System.Drawing.Point(12, 156);
            this.txtRespuesta.Name = "txtRespuesta";
            this.txtRespuesta.ReadOnly = true;
            this.txtRespuesta.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.txtRespuesta.Size = new System.Drawing.Size(460, 150);
            this.txtRespuesta.TabIndex = 2;
            this.txtRespuesta.Text = "";
            // 
            // lblRespuesta
            // 
            this.lblRespuesta.AutoSize = true;
            this.lblRespuesta.Location = new System.Drawing.Point(12, 140);
            this.lblRespuesta.Name = "lblRespuesta";
            this.lblRespuesta.Size = new System.Drawing.Size(109, 13);
            this.lblRespuesta.TabIndex = 5;
            this.lblRespuesta.Text = "Respuesta generada:";
            // 
            // btnGuardar
            // 
            this.btnGuardar.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnGuardar.Enabled = false;
            this.btnGuardar.Location = new System.Drawing.Point(316, 118);
            this.btnGuardar.Name = "btnGuardar";
            this.btnGuardar.Size = new System.Drawing.Size(75, 23);
            this.btnGuardar.TabIndex = 3;
            this.btnGuardar.Text = "Guardar BD";
            this.btnGuardar.UseVisualStyleBackColor = true;
            btnGuardar.Click += btnGuardar_Click;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.lblTokens});
            this.statusStrip1.Location = new System.Drawing.Point(0, 296);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(484, 22);
            this.statusStrip1.TabIndex = 4;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // lblTokens
            // 
            this.lblTokens.Name = "lblTokens";
            this.lblTokens.Size = new System.Drawing.Size(108, 17);
            this.lblTokens.Text = "Tokens utilizados: 0";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(484, 318);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.btnGuardar);
            this.Controls.Add(this.lblRespuesta);
            this.Controls.Add(this.txtRespuesta);
            this.Controls.Add(this.btnGenerar);
            this.Controls.Add(this.txtPregunta);
            this.MinimumSize = new System.Drawing.Size(500, 356);
            this.Name = "Form1";
            this.Text = "Generador de Documentos IA";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }



        #endregion
    }
}