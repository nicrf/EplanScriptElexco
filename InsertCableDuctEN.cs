// InsertCableDuct, Version 0.0.1, 2018-02-16
// By Nicolas Fournier nicolas.fournier@elexco.net
// Based on COmment script of Frank Schöneck
using System.Windows.Forms;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using System;
using System.Xml;
using System.IO;

//using System.Data;

public partial class frmInsertCableDuct : System.Windows.Forms.Form
{	
	private GroupBox groupBox1;
	private DataGridView dataGridView1;
	private TextBox textBoxDescription;
	private Button buttonOk;
	private Label label1;
	private TableLayoutPanel tableLayoutPanel1;
	private TextBox textBoxHeight;
	private Label label2;
	private DataGridViewTextBoxColumn NameColumn;
	private DataGridViewTextBoxColumn WidthColumn;
	private DataGridViewTextBoxColumn EnglishColumn;
	private DataGridViewTextBoxColumn EnvironnementColumn;
	private DataGridViewTextBoxColumn TextSizeColumn;
	private DataGridViewTextBoxColumn LineThicknessColumn;
	public string filename = "cableduct.xml";	

	#region Windows Form Designer generated code

	
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

	/// <summary>
	/// Required method for Designer support - do not modify
	/// the contents of this method with the code editor.
	/// </summary>
	private void InitializeComponent()
	{
		this.groupBox1 = new System.Windows.Forms.GroupBox();
		this.dataGridView1 = new System.Windows.Forms.DataGridView();
		this.textBoxDescription = new System.Windows.Forms.TextBox();
		this.buttonOk = new System.Windows.Forms.Button();
		this.label1 = new System.Windows.Forms.Label();
		this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
		this.textBoxHeight = new System.Windows.Forms.TextBox();
		this.label2 = new System.Windows.Forms.Label();
		this.NameColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.WidthColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.EnglishColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.EnvironnementColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.TextSizeColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.LineThicknessColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
		this.groupBox1.SuspendLayout();
		((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
		this.tableLayoutPanel1.SuspendLayout();
		this.SuspendLayout();
		// 
		// groupBox1
		// 
		this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
		| System.Windows.Forms.AnchorStyles.Left) 
		| System.Windows.Forms.AnchorStyles.Right)));
		this.groupBox1.Controls.Add(this.dataGridView1);
		this.groupBox1.Location = new System.Drawing.Point(3, 59);
		this.groupBox1.Name = "groupBox1";
		this.groupBox1.Size = new System.Drawing.Size(750, 286);
		this.groupBox1.TabIndex = 0;
		this.groupBox1.TabStop = false;
		this.groupBox1.Text = "Cable Duct";
		// 
		// dataGridView1
		// 
		this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
		this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
		this.NameColumn,
		this.WidthColumn,
		this.EnglishColumn,
		this.EnvironnementColumn,
		this.TextSizeColumn,
		this.LineThicknessColumn});
		this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
		this.dataGridView1.Location = new System.Drawing.Point(3, 16);
		this.dataGridView1.MultiSelect = false;
		this.dataGridView1.Name = "dataGridView1";
		this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
		this.dataGridView1.Size = new System.Drawing.Size(744, 267);
		this.dataGridView1.TabIndex = 0;
		this.dataGridView1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dataGridView1_EditingControlShowing);
		// 
		// textBoxDescription
		// 
		this.textBoxDescription.Dock = System.Windows.Forms.DockStyle.Fill;
		this.textBoxDescription.Location = new System.Drawing.Point(3, 23);
		this.textBoxDescription.Multiline = true;
		this.textBoxDescription.Name = "textBoxDescription";
		this.textBoxDescription.Size = new System.Drawing.Size(750, 30);
		this.textBoxDescription.TabIndex = 1;
		// 
		// buttonOk
		// 
		this.buttonOk.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
		| System.Windows.Forms.AnchorStyles.Left) 
		| System.Windows.Forms.AnchorStyles.Right)));
		this.buttonOk.Location = new System.Drawing.Point(3, 401);
		this.buttonOk.Name = "buttonOk";
		this.buttonOk.Size = new System.Drawing.Size(750, 25);
		this.buttonOk.TabIndex = 2;
		this.buttonOk.Text = "Create";
		this.buttonOk.UseVisualStyleBackColor = true;
		this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
		// 
		// label1
		// 
		this.label1.AutoSize = true;
		this.label1.Location = new System.Drawing.Point(3, 0);
		this.label1.Name = "label1";
		this.label1.Size = new System.Drawing.Size(133, 13);
		this.label1.TabIndex = 3;
		this.label1.Text = "Supplementary Description";
		this.label1.Click += new System.EventHandler(this.label1_Click);
		// 
		// tableLayoutPanel1
		// 
		this.tableLayoutPanel1.ColumnCount = 1;
		this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
		this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
		this.tableLayoutPanel1.Controls.Add(this.textBoxDescription, 0, 1);
		this.tableLayoutPanel1.Controls.Add(this.groupBox1, 0, 2);
		this.tableLayoutPanel1.Controls.Add(this.buttonOk, 0, 5);
		this.tableLayoutPanel1.Controls.Add(this.textBoxHeight, 0, 4);
		this.tableLayoutPanel1.Controls.Add(this.label2, 0, 3);
		this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
		this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
		this.tableLayoutPanel1.Name = "tableLayoutPanel1";
		this.tableLayoutPanel1.RowCount = 6;
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 11.11111F));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 88.88889F));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
		this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
		this.tableLayoutPanel1.Size = new System.Drawing.Size(756, 429);
		this.tableLayoutPanel1.TabIndex = 4;
		// 
		// textBoxHeight
		// 
		this.textBoxHeight.Dock = System.Windows.Forms.DockStyle.Fill;
		this.textBoxHeight.Location = new System.Drawing.Point(3, 371);
		this.textBoxHeight.Name = "textBoxHeight";
		this.textBoxHeight.Size = new System.Drawing.Size(750, 20);
		this.textBoxHeight.TabIndex = 4;
		this.textBoxHeight.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textBoxHeight_TextChanged);
		// 
		// label2
		// 
		this.label2.AutoSize = true;
		this.label2.Location = new System.Drawing.Point(3, 348);
		this.label2.Name = "label2";
		this.label2.Size = new System.Drawing.Size(44, 13);
		this.label2.TabIndex = 5;
		this.label2.Text = "Height :";
		// 
		// NameColumn
		// 
		this.NameColumn.DataPropertyName = "name";
		this.NameColumn.HeaderText = "Name";
		this.NameColumn.Name = "NameColumn";
		// 
		// WidthColumn
		// 
		this.WidthColumn.DataPropertyName = "width";
		this.WidthColumn.HeaderText = "Width (mm)";
		this.WidthColumn.Name = "WidthColumn";
		// 
		// EnglishColumn
		// 
		this.EnglishColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
		this.EnglishColumn.DataPropertyName = "eng";
		this.EnglishColumn.HeaderText = "Description";
		this.EnglishColumn.Name = "EnglishColumn";
		// 
		// EnvironnementColumn
		// 
		this.EnvironnementColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
		this.EnvironnementColumn.DataPropertyName = "env";
		this.EnvironnementColumn.HeaderText = "Environnement Description";
		this.EnvironnementColumn.Name = "EnvironnementColumn";
		// 
		// TextSizeColumn
		// 
		this.TextSizeColumn.DataPropertyName = "textsize";
		this.TextSizeColumn.HeaderText = "Text Size (mm)";
		this.TextSizeColumn.Name = "TextSizeColumn";
		// 
		// LineThicknessColumn
		// 
		this.LineThicknessColumn.DataPropertyName = "linethickness";
		this.LineThicknessColumn.HeaderText = "Line Thickness (mm)";
		this.LineThicknessColumn.Name = "LineThicknessColumn";
		// 
		// frmInsertCableDuct
		// 
		this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
		this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
		this.ClientSize = new System.Drawing.Size(756, 429);
		this.Controls.Add(this.tableLayoutPanel1);
		this.Name = "frmInsertCableDuct";
		this.Text = "Add Cable Duct 2D";
		this.groupBox1.ResumeLayout(false);
		((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
		this.tableLayoutPanel1.ResumeLayout(false);
		this.tableLayoutPanel1.PerformLayout();
		this.ResumeLayout(false);
	}        

	public class cableduct
	{
		public string name;
		public string width;
		public string eng;
		public string env;
		public string textsize;
		public string linethickness;
	}   

	public frmInsertCableDuct()
	{
		InitializeComponent();
		if (!File.Exists(filename))
		{ 
			using (var tw = new StreamWriter(filename, true, System.Text.Encoding.UTF8))
			{
				tw.WriteLine("<?xml version=\"1.0\" standalone=\"yes\"?>");
				tw.WriteLine("<dataset>");
				tw.WriteLine("<cableduct> ");
				tw.WriteLine("<name> 2X4\"</name>");
				tw.WriteLine("<width> 57.15 </width>");
				tw.WriteLine("<eng>CABLE DUCT 2X4</eng >");
				tw.WriteLine("<env>CANIVEAU 2X4</env>");
				tw.WriteLine("<textsize> 5.0 </textsize > ");
				tw.WriteLine("<linethickness> 5.00 </linethickness > ");
				tw.WriteLine("</cableduct> ");
				tw.WriteLine("</dataset> ");
				tw.Close();
			}
		}

		XmlTextReader xtr = new XmlTextReader(filename);
		// Create an XmlReader
		XmlDocument doc = new XmlDocument();
		doc.Load(filename);
		XmlNode node = doc.DocumentElement.SelectSingleNode("/dataset");
		foreach (XmlNode n in doc.DocumentElement.ChildNodes)
		{
			if (n.Name =="cableduct")
			{
				cableduct c = new cableduct();
				foreach (XmlNode val in n.ChildNodes)
				{
					switch(val.Name)
					{
						case "name":
							c.name = val.InnerText;
							break;
						case "width":
							c.width = val.InnerText;
							break;
						case "eng":
							c.eng = val.InnerText;
							break;
						case "env":
							c.env = val.InnerText;
							break;
						case "textsize":
							c.textsize = val.InnerText;
							break;
						case "linethickness":
							c.linethickness = val.InnerText;
							break;
					}
				}
				dataGridView1.Rows.Add(c.name, c.width, c.eng, c.env, c.textsize, c.linethickness);
			}
			string text = n.InnerText; //or loop through its children as well
		}
		//dataSet.ReadXml(filename);
		//dataGridView1.DataSource = dataSet.Tables[0];
	}
	#endregion


	//Menüpunkt anlegen
	[DeclareMenu()]
	public void CableDuct_Menu()
	{
		Eplan.EplApi.Gui.Menu oMenu = new Eplan.EplApi.Gui.Menu();
		oMenu.AddMenuItem("Add CableDuct", "DialogInsertCableDuct", "Submit CableDuct is executed", 35381, 0, false, false);
	}

	//Action um die Form aufzurufen
	[DeclareAction("DialogInsertCableDuct")]
	public void DialogInsertCableDuct_action()
	{
		frmInsertCableDuct frm = new frmInsertCableDuct();
		frm.ShowDialog();
		return;
	}

	private void SaveXML()
	{
		using (XmlWriter writer = XmlWriter.Create(filename))
		{
			writer.WriteStartDocument();
			writer.WriteStartElement("dataset");
			for (int i = 0; i < dataGridView1.Rows.Count - 1; i++) // -1 = empty rows
			{
				writer.WriteStartElement("cableduct");

				for (int j = 0; j < dataGridView1.Columns.Count; j++)
				{
					 string value;
					if (dataGridView1.Rows[i].Cells[j].Value ==null)
					 	value = "";
					else
						value = dataGridView1.Rows[i].Cells[j].Value.ToString();
					 
					writer.WriteElementString(dataGridView1.Columns[j].DataPropertyName,value);// dataGridView1.Rows[i].Cells[j].Value?.ToString());
				}               
				writer.WriteEndElement();
			}
			writer.WriteEndElement();
			writer.WriteEndDocument();
		}
	}
	private void buttonOk_Click(object sender, EventArgs e)
	{
		//Save Data
		//dataSet.Tables[0].WriteXml(filename);
		SaveXML();

		//Create Macro
		if ((dataGridView1.SelectedRows.Count > 0) && (!string.IsNullOrWhiteSpace(textBoxHeight.Text)))
		{   
				CreateCableDuct();
		}
	}

	private void textBoxHeight_TextChanged(object sender, KeyPressEventArgs e)
	{
		if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&   (e.KeyChar != '.'))
		{
			e.Handled = true;
		}
	}

	private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
	{
		e.Control.KeyPress -= new KeyPressEventHandler(textBoxHeight_TextChanged);
		if (dataGridView1.CurrentCell.ColumnIndex == dataGridView1.Columns["WidthColumn"].Index ||
			dataGridView1.CurrentCell.ColumnIndex == dataGridView1.Columns["TextSizeColumn"].Index ||
			dataGridView1.CurrentCell.ColumnIndex == dataGridView1.Columns["LineThicknessColumn"].Index )
		{
			TextBox tb = e.Control as TextBox;
			if (tb != null)
			{
				tb.KeyPress += new KeyPressEventHandler(textBoxHeight_TextChanged);
			}
		}
	}

	//Button OK Click
	private void CreateCableDuct()
	{
		string sbuf;
		//Level of comment (A411 = 519(EPLAN519, Graphic.Comment))
		string sEbene = "519";

		//Linetype of the comment(A412 = L(Layer) / 0(pulled through) / 41(~~~~~))
		string sLinientyp = "0";

		//Pattern length of the comment (A415 = L(Layer) / -1.5(1,50 mm) / -32(32,00 mm))
		string sPatternLenght = "-1.5";

		//Color of the comment (A413 = 0 (black) / 1 (red) / 2 (yellow) / 3 (light green) / 4 (light blue) / 5 (dark blue) / 6 (violet) / 8 (white) / 40 (orange) / -4 Watermark / -5 Inverse to background)
		string sFarbe = "-4";

        //Fill the surface
        string sFill = "1";
		double dStartX = 64;
		double dStartY = 248;
		string sAngle = "";

		string sLineThickness = "2";
		if (dataGridView1.SelectedRows[0].Cells["LineThicknessColumn"].Value !=null)
			sLineThickness = dataGridView1.SelectedRows[0].Cells["LineThicknessColumn"].Value.ToString();

		//A966 = Alignment 0 = Layer, 4 = Middle Left, 5 = Middle centre
		string sAlignment = "5";

		//A503 = Centered , 256 = All , 128 = First text
		string sCentered = "256";

		//Text size
		string sTexteSize = "5";
		if (dataGridView1.SelectedRows[0].Cells["TextSizeColumn"].Value !=null)
			sTexteSize = dataGridView1.SelectedRows[0].Cells["TextSizeColumn"].Value.ToString();


		//Size	
		
		if (dataGridView1.SelectedRows[0].Cells["WidthColumn"].Value !=null)
			sbuf = dataGridView1.SelectedRows[0].Cells["WidthColumn"].Value.ToString();
		else
			return;
		
		double dWidth = double.Parse(sbuf);	
		double dHeight =  double.Parse(textBoxHeight.Text);

		//CableDuctText (A511)
		string sCableDuctText = "";
		string sCableDuctTextEnglish,sCableDuctTextEnv;
		
		if (dataGridView1.SelectedRows[0].Cells["EnglishColumn"].Value !=null)
			sCableDuctTextEnglish = dataGridView1.SelectedRows[0].Cells["EnglishColumn"].Value.ToString();
		else
			return;

		if (dataGridView1.SelectedRows[0].Cells["EnvironnementColumn"].Value !=null)
			sCableDuctTextEnv = dataGridView1.SelectedRows[0].Cells["EnvironnementColumn"].Value.ToString();
		else
			return;
			
		if (!string.IsNullOrWhiteSpace(textBoxDescription.Text))
		{
			sCableDuctTextEnglish = sCableDuctTextEnglish  + " "  + textBoxDescription.Text; //
			sCableDuctTextEnv = sCableDuctTextEnv  + " "  + textBoxDescription.Text; //
		}
		sCableDuctText = "en_US@"+ sCableDuctTextEnglish + ";fr_FR@" + sCableDuctTextEnv + ";";//??_??@" + sCableDuctTextEnv + ";";

		/*if (sCableDuctText.EndsWith(Environment.NewLine)) //Kommentar darf nicht mit Zeilenumbruch enden
		{ sCableDuctText = sCableDuctText.Substring(0, sCableDuctText.Length - 2); }
		sCableDuctText = sCableDuctText.Replace(Environment.NewLine, "&#10;"); //Kommentar Zeilenumbruch umwandeln
		sCableDuctText = "??_??@" + sCableDuctText; //Kommentar MultiLanguage String
		if (!sCableDuctText.EndsWith(";")) //Kommentar muss mit ";" enden
		{ sCableDuctText += ";"; }*/

		//Pfad und Dateiname der Temp.datei
		string sTempFile;
#if DEBUG
		sTempFile = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\tmpInsertCableDuct.ema";
#else
		sTempFile = PathMap.SubstitutePath(@"$(TMP)") + @"\tmpInsertCableDuct.ema";
#endif
		XmlWriterSettings settings = new XmlWriterSettings();
		settings.Indent = true;
		XmlWriter writer = XmlWriter.Create(sTempFile, settings);

		//Macro file content
		writer.WriteRaw("\n<EplanPxfRoot Name=\"#Kommentar\" Type=\"WindowMacro\" Version=\"2.2.6360\" PxfVersion=\"1.23\" SchemaVersion=\"1.7.6360\" Source=\"\" SourceProject=\"\" Description=\"\" ConfigurationFlags=\"0\" NumMainObjects=\"0\" NumProjectSteps=\"0\" NumMDSteps=\"0\" Custompagescaleused=\"true\" StreamSchema=\"EBitmap,BaseTypes,2,1,2;EPosition3D,BaseTypes,0,3,2;ERay3D,BaseTypes,0,4,2;EStreamableVector,BaseTypes,0,5,2;DMNCDataSet,TrDMProject,1,20,2;DMNCDataSetVector,TrDMProject,1,21,2;DMPlaceHolderRuleData,TrDMProject,0,22,2;Arc3d@W3D,W3dBaseGeometry,0,36,2;Box3d@W3D,W3dBaseGeometry,0,37,2;Circle3d@W3D,W3dBaseGeometry,0,38,2;Color@W3D,W3dBaseGeometry,0,39,2;ContourPlane3d@W3D,W3dBaseGeometry,0,40,2;CTexture@W3D,W3dBaseGeometry,0,41,2;CTextureMap@W3D,W3dBaseGeometry,0,42,2;Line3d@W3D,W3dBaseGeometry,0,43,2;Linetype@W3D,W3dBaseGeometry,0,44,2;Material@W3D,W3dBaseGeometry,3,45,2;Path3d@W3D,W3dBaseGeometry,0,46,2;Mesh3dX@W3D,W3dMeshModeller,2,47,2;MeshBox@W3D,W3dMeshModeller,5,48,2;MeshMate@W3D,W3dMeshModeller,7,49,2;MeshMateFace@W3D,W3dMeshModeller,1,50,2;MeshMateGrid@W3D,W3dMeshModeller,8,51,2;MeshMateGridLine@W3D,W3dMeshModeller,1,52,2;MeshMateLine@W3D,W3dMeshModeller,1,53,2;MeshText3dX@W3D,W3dMeshModeller,0,55,2;BaseTextLine@W3D,W3dMeshModeller,2,56,2;Mesh3d@W3D,W3dMeshModeller,8,57,2;MeshEdge3d@W3D,W3dMeshModeller,0,58,2;MeshFace3d@W3D,W3dMeshModeller,2,59,2;MeshPoint3d@W3D,W3dMeshModeller,1,60,2;MeshPolygon3d@W3D,W3dMeshModeller,1,61,2;MeshSimpleTextureTriangle3d@W3D,W3dMeshModeller,2,62,2;MeshSimpleTriangle3d@W3D,W3dMeshModeller,1,63,2;MeshTriangle3d@W3D,W3dMeshModeller,2,64,2;MeshTriangleFace3d@W3D,W3dMeshModeller,0,65,2;MeshTriangleFaceEdge3D@W3D,W3dMeshModeller,0,66,2\">");
		writer.WriteRaw("\n <MacroVariant MacroFuncType=\"1\" VariantId=\"0\" ReferencePoint=\"64/248/0\" Version=\"2.2.6360\" PxfVersion=\"1.23\" SchemaVersion=\"1.7.6360\" Source=\"\" SourceProject=\"\" Description=\"\" ConfigurationFlags=\"0\" DocumentType=\"1\" Customgost=\"0\">");
		writer.WriteRaw("\n  <O4 Build=\"6360\" A1=\"4/18\" A3=\"0\" A13=\"0\" A14=\"0\" A47=\"1\" A48=\"1362057551\" A50=\"1\" A59=\"1\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A431=\"1\" A1101=\"17\" A1102=\"\" A1103=\"\">");

		//If you want to group
		writer.WriteRaw("\n  <O26 Build=\"6360\" A1=\"26/128740\" A3=\"0\" A13=\"0\" A14=\"0\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A431=\"1\">");

		//Properties CableDuct
		//Rectangle Property
		//A1656 = fill the surface
		//A1655 = radius
		//A1654 = round corner
		//A1653  =  angle
		//A1651 = x / y statt
		//A1652 = x / y end
		//A406 = invisible
		//A411 = Layer
		//A414 = Line thickness
		//A415 =  Pattern length
		//A416 = Line end style (1 = Rectangle)
		//A1657

		//Wateramrk rectangle
		sFarbe = "-4"; // Watermark
		//sPatternLenght = "L";
		writer.WriteRaw("\n  <O89 Build=\"9661\"  A1=\"89/339\"  A3=\"0\"  A13=\"0\"  A14=\"0\"  A404=\"1\"  A405=\"64\"  A406=\"0\"  A407=\"0\"  A411=\"100\"  A412=\"" + sLinientyp + "\"  A413=\"" + sFarbe +
		 "\" A414=\"" + sLineThickness + "\"  A415=\"" + sPatternLenght + "\"  A416=\"0\"  A1651=\"64/248\"  A1652=\""+  Convert.ToString(dWidth + dStartX) + "/" + Convert.ToString(dHeight + dStartY)+ "\"  A1653=\"0\"  A1654=\"0\"  A1655=\"0\"   A1656=\""+sFill+"\"   A1657=\" 0\" />");
		//Border  rectangle
		//sPatternLenght = "-10.0";
		sFarbe = "-5"; //Inverse background
		//sLineThickness = "5";
        writer.WriteRaw("\n  <O89 Build=\"9661\"  A1=\"89/339\"  A3=\"0\"  A13=\"0\"  A14=\"0\"  A404=\"1\"  A405=\"64\"  A406=\"0\"  A407=\"0\"  A411=\"100\"  A412=\"" + sLinientyp + "\"  A413=\"" + sFarbe +
		 "\" A414=\"" + sLineThickness + "\"   A415=\"" + sPatternLenght + "\"  A416=\"0\"  A1651=\"64/248\"  A1652=\""+ Convert.ToString(dWidth + dStartX) + "/" + Convert.ToString(dHeight + dStartY) + "\"  A1653=\"0\"  A1654=\"0\"  A1655=\"0\" A1656=\"0\"  A1657=\" 0\" />");


		//Properties text		
		//A411 = Layer
		//A413 = Color
		//A503 = Centered , 256 = All , 128 = First text
		//A501 = Ref position
		//A966 = Alignment 0 = Layer, 4 = Middle Left, 5 = Middle centre
		//A961 = Text size
		//A962 = Angle (radian), L = Layer
		//A964 = Color too		
		//A511 = Texte ="en_US@angalsi;fr_FR@francais;"
		//sCableDuctText = "en_US@angalsi;fr_FR@francais;";
		double dAngle = 90 * 3.1416 / 180;
		sAngle = dAngle.ToString();
		writer.WriteRaw("\n  <O30 Build=\"9661\" A1=\"30/11504\" A3=\"0\" A13=\"0\" A14=\"0\" A404=\"33\" A405=\"64\" A406=\"0\" A407=\"0\" A411=\"108\" A412=\"L\" A413=\"-5\" A414=\"L\" A415=\"L\" A416=\"0\" A501=\""+ Convert.ToString(dWidth/2+ dStartX) +"/" + Convert.ToString(dHeight/2+ dStartY) +"\" A503=\""+sCentered +
		"\" A506=\"0\" A511=\"" + sCableDuctText + "\">");
		writer.WriteRaw("\n <S54x505 A961=\""+ sTexteSize +"\" A962=\""+ sAngle +"\" A963=\"0\" A964=\"-5\" A965=\"16384\" A966=\""+sAlignment+"\" A967=\"8.48\" A968=\"3.75\" A969=\"0\" A4000=\"L\" A4001=\"L\" A4013=\"0\"/>");
		writer.WriteRaw("\n  </O30>");		


		//If you want to group
		writer.WriteRaw("\n  </O26>");
		
		writer.WriteRaw("\n  <O37 Build=\"6360\" A1=\"37/128743\" A3=\"1\" A13=\"0\" A14=\"0\" A404=\"1\" A405=\"64\" A406=\"0\" A407=\"0\" A682=\"1\" A683=\"26/128740\" A684=\"0\" A687=\"8\" A688=\"2\" A689=\"-1\" A690=\"-1\" A691=\"0\" A693=\"1\" A792=\"0\" A793=\"0\" A794=\"0\" A1261=\"0\" A1262=\"44\" A1263=\"0\" A1631=\"0/-6.34911032028501\" A1632=\"8/0\">");
		writer.WriteRaw("\n  <S109x692 Build=\"6360\" A3=\"0\" A13=\"0\" A14=\"0\" R1906=\"165/128741\"/>");
		writer.WriteRaw("\n  <S109x692 Build=\"6360\" A3=\"0\" A13=\"0\" A14=\"0\" R1906=\"30/128742\"/>");
		writer.WriteRaw("\n  <S40x1201 A762=\"64/254.349110320285\">");
		writer.WriteRaw("\n  <S39x761 A751=\"1\" A752=\"0\" A753=\"0\" A754=\"1\"/>");
		writer.WriteRaw("\n  </S40x1201>");
		writer.WriteRaw("\n  <S89x5 Build=\"6360\" A3=\"0\" A4=\"1\" R7=\"37/128743\" A13=\"0\" A14=\"0\" A404=\"9\" A405=\"64\" A406=\"0\" A407=\"0\" A411=\"308\" A412=\"L\" A413=\"L\" A414=\"L\" A415=\"L\" A416=\"0\" A1651=\"0/-6.34911032028501\" A1652=\"8/0\" A1653=\"0\" A1654=\"0\" A1655=\"0\" A1656=\"0\" A1657=\"0\"/>");
		writer.WriteRaw("\n  </O37>");
		writer.WriteRaw("\n  </O4>");
		writer.WriteRaw("\n </MacroVariant>");
		writer.WriteRaw("\n</EplanPxfRoot>");

		// Write the XML to file and close the writer.
		writer.Flush();
		writer.Close();

		//Makro einfügen
#if DEBUG
		MessageBox.Show(sTempFile);
#else
		CommandLineInterpreter oCli = new CommandLineInterpreter();
		ActionCallingContext oAcc = new ActionCallingContext();
		oAcc.AddParameter("Name", "XMIaInsertMacro");
		oAcc.AddParameter("filename", sTempFile);
		oAcc.AddParameter("variant", "0");
		oCli.Execute("XGedStartInteractionAction", oAcc);
#endif

		//Beenden
		Close();
		return;
	}

	 private void label1_Click(object sender, EventArgs e)
	{
            System.Diagnostics.Process.Start(filename);
	}

}