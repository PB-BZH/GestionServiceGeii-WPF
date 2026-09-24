/*
╔════════════════════════════════════════════════════════════════════════════════╗
║                                                                                ║
║                    ───────────────────────────────────                         ║
║                      © Copyright PB-BZH Concept 2026                           ║
║                    ───────────────────────────────────                         ║
║                                                                                ║
║                 contact : mailto:admin@pb-bzh-concept.fr                       ║
╚════════════════════════════════════════════════════════════════════════════════╝

╔════════════════════════════════════════════════════════════════════════════════╗
║  Auteur : Patrick Bourges - PB-BZH Concept                                     ║
║  Le 17/6/2026 - 11:30
╟────────────────────────────────────────────────────────────────────────────────║
║     Projet Visual Studio Professional 2026 : GestionServicesGEII
╟────────────────────────────────────────────────────────────────────────────────║
║     Version : 1.5.3
╟────────────────────────────────────────────────────────────────────────────────║
║                Visual Studio Professional 2026 - Insiders                      ║
║                ──────────────────────────────────────────                      ║
║  Langage     : C# 14                                                           ║
║  Technologie : .NET 10 / WinForms                                              ║
║  Plateforme  : Windows 10 / Windows 11                                         ║
║  Encodage    : UTF-8                                                           ║
╟────────────────────────────────────────────────────────────────────────────────║
║  Nom de fichier : FichierRegistre.Designer.cs
╚════════════════════════════════════════════════════════════════════════════════╝
*/
namespace GestionServicesGEII.Shared.Librairie_Registre
{
    partial class FichierRegistre
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
      this.TexteRépertoireDeTravail = new System.Windows.Forms.Label();
      this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
      this.titreFenetre = new System.Windows.Forms.TextBox();
      this.textBox1 = new System.Windows.Forms.TextBox();
      this.répertoireDeTravail = new System.Windows.Forms.Label();
      this.tableLayoutPanel1.SuspendLayout();
      this.SuspendLayout();
      // 
      // TexteRépertoireDeTravail
      // 
      this.TexteRépertoireDeTravail.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
      this.TexteRépertoireDeTravail.AutoSize = true;
      this.TexteRépertoireDeTravail.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.TexteRépertoireDeTravail.Location = new System.Drawing.Point(3, 71);
      this.TexteRépertoireDeTravail.Margin = new System.Windows.Forms.Padding(3);
      this.TexteRépertoireDeTravail.Name = "TexteRépertoireDeTravail";
      this.TexteRépertoireDeTravail.Size = new System.Drawing.Size(136, 47);
      this.TexteRépertoireDeTravail.TabIndex = 1;
      this.TexteRépertoireDeTravail.Text = "Répertoire de travail :";
      this.TexteRépertoireDeTravail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
      // 
      // tableLayoutPanel1
      // 
      this.tableLayoutPanel1.ColumnCount = 3;
      this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
      this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 331F));
      this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 106F));
      this.tableLayoutPanel1.Controls.Add(this.titreFenetre, 0, 0);
      this.tableLayoutPanel1.Controls.Add(this.textBox1, 1, 3);
      this.tableLayoutPanel1.Controls.Add(this.TexteRépertoireDeTravail, 0, 2);
      this.tableLayoutPanel1.Controls.Add(this.répertoireDeTravail, 1, 2);
      this.tableLayoutPanel1.Location = new System.Drawing.Point(1, -1);
      this.tableLayoutPanel1.Name = "tableLayoutPanel1";
      this.tableLayoutPanel1.RowCount = 4;
      this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 45.58823F));
      this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 54.41177F));
      this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 53F));
      this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 154F));
      this.tableLayoutPanel1.Size = new System.Drawing.Size(579, 276);
      this.tableLayoutPanel1.TabIndex = 4;
      // 
      // titreFenetre
      // 
      this.titreFenetre.BackColor = System.Drawing.SystemColors.MenuHighlight;
      this.tableLayoutPanel1.SetColumnSpan(this.titreFenetre, 3);
      this.titreFenetre.Cursor = System.Windows.Forms.Cursors.Arrow;
      this.titreFenetre.Dock = System.Windows.Forms.DockStyle.Fill;
      this.titreFenetre.Font = new System.Drawing.Font("Cambria", 12F, ((System.Drawing.FontStyle)((System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic))));
      this.titreFenetre.ForeColor = System.Drawing.Color.Gold;
      this.titreFenetre.Location = new System.Drawing.Point(0, 0);
      this.titreFenetre.Margin = new System.Windows.Forms.Padding(0);
      this.titreFenetre.Name = "titreFenetre";
      this.titreFenetre.ReadOnly = true;
      this.titreFenetre.Size = new System.Drawing.Size(579, 26);
      this.titreFenetre.TabIndex = 4;
      this.titreFenetre.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
      // 
      // textBox1
      // 
      this.textBox1.Location = new System.Drawing.Point(145, 124);
      this.textBox1.Name = "textBox1";
      this.textBox1.Size = new System.Drawing.Size(325, 20);
      this.textBox1.TabIndex = 2;
      // 
      // répertoireDeTravail
      // 
      this.répertoireDeTravail.AccessibleName = string.Empty;
      this.répertoireDeTravail.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
      this.répertoireDeTravail.AutoSize = true;
      this.répertoireDeTravail.Font = new System.Drawing.Font("Times New Roman", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.répertoireDeTravail.Location = new System.Drawing.Point(145, 71);
      this.répertoireDeTravail.Margin = new System.Windows.Forms.Padding(3);
      this.répertoireDeTravail.Name = "répertoireDeTravail";
      this.répertoireDeTravail.Size = new System.Drawing.Size(61, 47);
      this.répertoireDeTravail.TabIndex = 3;
      this.répertoireDeTravail.Text = "non défini";
      this.répertoireDeTravail.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
      // 
      // FichierRegistre
      // 
      this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
      this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
      this.ClientSize = new System.Drawing.Size(579, 275);
      this.Controls.Add(this.tableLayoutPanel1);
      this.Name = "FichierRegistre";
      this.Text = "Sauvegarde des paramètres";
      this.tableLayoutPanel1.ResumeLayout(false);
      this.tableLayoutPanel1.PerformLayout();
      this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label TexteRépertoireDeTravail;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label répertoireDeTravail;
        private System.Windows.Forms.TextBox textBox1;
        public System.Windows.Forms.TextBox titreFenetre;
    }
}