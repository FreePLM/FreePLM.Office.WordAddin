using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using FreePLM.Office.WordAddin.Models;

namespace FreePLM.Office.WordAddin.Forms
{
    public class SearchDocumentForm : Form
    {
        private TextBox txtObjectId;
        private TextBox txtFileName;
        private TextBox txtProject;
        private TextBox txtOwner;
        private ComboBox cmbStatus;
        private Button btnSearch;
        private Button btnClear;
        private DataGridView dgvResults;
        private Button btnOpen;
        private Button btnCancel;

        public string SelectedObjectId { get; private set; }
        public List<DocumentSearchResultDto> SearchResults { get; private set; }

        public event EventHandler<SearchRequestEventArgs> SearchRequested;

        public SearchDocumentForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Search PLM Documents";
            this.Width = 900;
            this.Height = 650;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimizeBox = false;
            this.MaximizeBox = true;

            // Search criteria panel
            var panelSearch = new Panel
            {
                Location = new Point(10, 10),
                Size = new Size(860, 150),
                BorderStyle = BorderStyle.FixedSingle
            };

            // ObjectId
            var lblObjectId = new Label { Text = "Object ID:", Location = new Point(10, 15), Width = 100 };
            txtObjectId = new TextBox { Location = new Point(120, 12), Width = 200 };

            // FileName
            var lblFileName = new Label { Text = "File Name:", Location = new Point(340, 15), Width = 100 };
            txtFileName = new TextBox { Location = new Point(450, 12), Width = 200 };

            // Project
            var lblProject = new Label { Text = "Project:", Location = new Point(10, 50), Width = 100 };
            txtProject = new TextBox { Location = new Point(120, 47), Width = 200 };

            // Owner
            var lblOwner = new Label { Text = "Owner:", Location = new Point(340, 50), Width = 100 };
            txtOwner = new TextBox { Location = new Point(450, 47), Width = 200 };

            // Status
            var lblStatus = new Label { Text = "Status:", Location = new Point(10, 85), Width = 100 };
            cmbStatus = new ComboBox
            {
                Location = new Point(120, 82),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbStatus.Items.AddRange(new object[] { "", "Private", "InWork", "Frozen", "Released", "Obsolete" });
            cmbStatus.SelectedIndex = 0;

            // Buttons
            btnSearch = new Button { Text = "Search", Location = new Point(450, 80), Width = 100 };
            btnSearch.Click += BtnSearch_Click;

            btnClear = new Button { Text = "Clear", Location = new Point(560, 80), Width = 100 };
            btnClear.Click += BtnClear_Click;

            panelSearch.Controls.AddRange(new Control[]
            {
                lblObjectId, txtObjectId, lblFileName, txtFileName,
                lblProject, txtProject, lblOwner, txtOwner,
                lblStatus, cmbStatus, btnSearch, btnClear
            });

            // Results grid
            dgvResults = new DataGridView
            {
                Location = new Point(10, 170),
                Size = new Size(860, 380),
                AutoGenerateColumns = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };

            dgvResults.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { HeaderText = "Object ID", DataPropertyName = "ObjectId", Width = 140 },
                new DataGridViewTextBoxColumn { HeaderText = "File Name", DataPropertyName = "FileName", Width = 180 },
                new DataGridViewTextBoxColumn { HeaderText = "Revision", DataPropertyName = "CurrentRevision", Width = 80 },
                new DataGridViewTextBoxColumn { HeaderText = "Status", DataPropertyName = "Status", Width = 90 },
                new DataGridViewTextBoxColumn { HeaderText = "Owner", DataPropertyName = "Owner", Width = 120 },
                new DataGridViewTextBoxColumn { HeaderText = "Project", DataPropertyName = "Project", Width = 100 },
                new DataGridViewCheckBoxColumn { HeaderText = "Checked Out", DataPropertyName = "IsCheckedOut", Width = 90 }
            });

            dgvResults.DoubleClick += DgvResults_DoubleClick;

            // Bottom buttons
            btnOpen = new Button { Text = "Open Selected", Location = new Point(670, 560), Width = 100 };
            btnOpen.Click += BtnOpen_Click;

            btnCancel = new Button { Text = "Cancel", Location = new Point(780, 560), Width = 90 };
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Click += (s, e) => this.Close();

            this.Controls.AddRange(new Control[]
            {
                panelSearch, dgvResults, btnOpen, btnCancel
            });

            this.CancelButton = btnCancel;
            this.AcceptButton = btnSearch;
        }

        private void BtnSearch_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtObjectId.Text) &&
                string.IsNullOrWhiteSpace(txtFileName.Text) &&
                string.IsNullOrWhiteSpace(txtProject.Text) &&
                string.IsNullOrWhiteSpace(txtOwner.Text) &&
                string.IsNullOrWhiteSpace(cmbStatus.Text))
            {
                MessageBox.Show("Please enter at least one search criteria.", "Search",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var args = new SearchRequestEventArgs
            {
                ObjectId = txtObjectId.Text.Trim(),
                FileName = txtFileName.Text.Trim(),
                Project = txtProject.Text.Trim(),
                Owner = txtOwner.Text.Trim(),
                Status = cmbStatus.Text
            };

            SearchRequested?.Invoke(this, args);
        }

        private void BtnClear_Click(object sender, EventArgs e)
        {
            txtObjectId.Clear();
            txtFileName.Clear();
            txtProject.Clear();
            txtOwner.Clear();
            cmbStatus.SelectedIndex = 0;
            dgvResults.DataSource = null;
        }

        private void DgvResults_DoubleClick(object sender, EventArgs e)
        {
            BtnOpen_Click(sender, e);
        }

        private void BtnOpen_Click(object sender, EventArgs e)
        {
            if (dgvResults.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a document to open.", "No Selection",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selectedDoc = dgvResults.SelectedRows[0].DataBoundItem as DocumentSearchResultDto;
            if (selectedDoc != null)
            {
                SelectedObjectId = selectedDoc.ObjectId;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        public void SetSearchResults(List<DocumentSearchResultDto> results)
        {
            SearchResults = results;
            dgvResults.DataSource = results;

            if (results == null || results.Count == 0)
            {
                MessageBox.Show("No documents found.", "Search Results",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }

    public class SearchRequestEventArgs : EventArgs
    {
        public string ObjectId { get; set; }
        public string FileName { get; set; }
        public string Project { get; set; }
        public string Owner { get; set; }
        public string Status { get; set; }
    }
}
