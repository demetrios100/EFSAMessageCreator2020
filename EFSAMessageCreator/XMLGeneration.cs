// <copyright file="XMLGeneration.cs" company="EFSAUsersGroup">Copyright (c) EFSA Users Group. All rights reserved.</copyright>
// <author>Demetrios Ioannides</author>
// <email>dvi1@columbia.edu</email>
// <summary>Generating XML files for Submission to EFSA</summary>

namespace EFSAMessageCreator
{
    #region using statements
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Threading;
    using System.Xml;
    using System.Xml.Schema;
    using System.Xml.Serialization;
    using Microsoft.VisualBasic.FileIO;
    #endregion

    /// <summary>
    /// The Class used for XML Generation
    /// </summary>
    public class XMLGeneration
    {
        #region Variables
        /// <summary>
        /// The Main Window Object
        /// </summary>
        private MainWindow mainWindowObject;

        /// <summary>
        /// The Full Name of the embedded XSD File
        /// </summary>
        private string xsdFileFullName;

        /// <summary>
        /// The Full Name of the embedded Element Mapping File
        /// </summary>
        private string elementMappingFileFullName;

        /// <summary>
        /// The Name of the Output XML File
        /// </summary>
        private string outputXMLFileName;

        /// <summary>
        /// The Validation Error Count
        /// </summary>
        private int validationErrorCount = 0;

        /// <summary>
        /// The Path of the Data File
        /// </summary>
        private string dataFilePath;

        /// <summary>
        /// Delimiter Type
        /// </summary>
        private DelimiterType delimiterType;

        #endregion

        #region Constructor
        /// <summary>
        /// Initializes a new instance of the <see cref="XMLGeneration"/> class
        /// </summary>
        /// <param name="mainWindowObject">The Object of the Application's Main Window</param>
        /// <param name="xsdFileFullName">The Full Name of the embedded XSD File</param>
        /// <param name="elementMappingFileFullName">The Full Name of the embedded Element Mapping File</param>
        /// <param name="outputXMLFileName">The Name of the Output XML File</param>
        /// <param name="outputXMLEncoding">The Encoding (UTF-8 or Unicode UTF-16) of the Output XML File</param>
        /// <param name="fileType">The File Type</param>
        /// <param name="delimiterType">The Delimiter Type</param>
        /// <param name="dataFilePath">The Selected Data File Path</param>
        public XMLGeneration(
            MainWindow mainWindowObject,
            string xsdFileFullName, 
            string elementMappingFileFullName, 
            string outputXMLFileName, 
            System.Text.Encoding outputXMLEncoding,
            FileType fileType,
            DelimiterType delimiterType,
            string dataFilePath)
        {
            try
            {
                this.mainWindowObject = mainWindowObject;
                this.outputXMLFileName = outputXMLFileName;
                this.xsdFileFullName = xsdFileFullName;
                this.elementMappingFileFullName = elementMappingFileFullName;
                this.dataFilePath = dataFilePath;
                this.delimiterType = delimiterType;

                // Get the Data Rows
                List<DataRow> rows;
                switch (fileType)
                {
                    case FileType.DBF:
                        rows = this.GetDBFDataRows();
                        this.SerializeRows(rows, outputXMLEncoding);
                        break;
                    case FileType.Delimited:
                        rows = this.GetDelimitedFileDataRows();
                        this.SerializeRows(rows, outputXMLEncoding);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region Refresh Controls
        /// <summary>
        /// Do Events (refresh controls)
        /// </summary>
        private static void DoEvents()
        {
            try
            {
                Application.Current.Dispatcher.Invoke(
                    DispatcherPriority.Background,
                    new Action(
                        delegate { }));
            }
            catch
            {
            }
        }
        #endregion

        #region Serialize Data Rows
        /// <summary>
        /// Serialize Data Rows
        /// </summary>
        /// <param name="rows">The List of Data Rows to be Serialized</param>
        /// <param name="outputXMLEncoding">The XML Encoding of the output</param>
        private void SerializeRows(List<DataRow> rows, System.Text.Encoding outputXMLEncoding)
        {
            try
            {
                // Limit the number of Rows as specified
                if (rows != null)
                {
                    rows = this.LimitRows(rows);

                    // Get the Element Mapping Table and perform the Matching Test
                    DataTable elementMappingTable = this.ReadElementMapping();
                    this.PerformMatchingTest(elementMappingTable);

                    // Get the results
                    List<recordType> results = this.GetResults(rows, elementMappingTable);

                    if (results != null)
                    {
                        // Create the Message object
                        message messageObject = new message();
                        messageObject.payload = new payload();
                        messageObject.payload.dataset = results.ToArray();

                        // Serialize XML
                        if (this.Serialize(messageObject, outputXMLEncoding) == true)
                        {
                            // Validate the produced XML file
                            this.Validate();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
        
        #region Connect to the dbf file and get the Data Rows
        /// <summary>
        /// Connect to the DBF file and get the Data Rows
        /// </summary>
        /// <returns>The list of Data Rows</returns>
        private List<DataRow> GetDBFDataRows()
        {
            try
            {
                this.mainWindowObject.StatusLabel.Content = "Reading the dbf file...";
                DoEvents();

                FileInfo fileInfoDataFile = new FileInfo(this.dataFilePath);
                string connectionString = @"Provider=VFPOLEDB.1;Data Source=" + this.dataFilePath;
                OleDbConnection conn = new OleDbConnection(connectionString);
                OleDbDataAdapter da = new OleDbDataAdapter("select * from " + fileInfoDataFile.Name.ToLower().Replace(".dbf", string.Empty), conn);

                DataSet dbfDs = new DataSet();
                da.Fill(dbfDs);
                conn.Close();
                conn.Dispose();
                da.Dispose();
                List<DataRow> rows = (from r in dbfDs.Tables[0].AsEnumerable()
                                      select r).ToList();
                return rows;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data File Reading problem: " + ex.Message);
                return null;
            }
        }
        #endregion
        
        #region Connect to the Delimited file and get the Data Rows
        /// <summary>
        /// Connect to the Delimited file and get the Data Rows
        /// </summary>
        /// <returns>The list of Data Rows</returns>
        private List<DataRow> GetDelimitedFileDataRows()
        {
            try
            {
                this.mainWindowObject.StatusLabel.Content = "Reading the Delimited file...";
                DoEvents();

                DataTable dt = this.PopulateDataTableFromDelimitedFile(this.dataFilePath);

                List<DataRow> rows = (from r in dt.AsEnumerable()
                                      select r).ToList();
                return rows;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Data File Reading problem: " + ex.Message);
                return null;
            }
        }
        #endregion
        
        #region Limit to the number specified
        /// <summary>
        /// Limit rows to the number specified
        /// </summary>
        /// <param name="rows">The complete list of rows</param>
        /// <returns>The list of rows limited by the specified number</returns>
        private List<DataRow> LimitRows(List<DataRow> rows)
        {
            int limit = 0;
            if (int.TryParse(this.mainWindowObject.LimitTextBox.Text, out limit))
            {
                rows = rows.Take(limit).ToList();
                this.outputXMLFileName = this.outputXMLFileName.Replace(".xml", "_First_" + limit.ToString() + ".xml");
            }

            return rows;
        }
        #endregion
        
        #region Read the contents of ElementMapping XML into "elementMappingTable"
        /// <summary>
        /// Read the contents of ElementMapping XML into "elementMappingTable"
        /// </summary>
        /// <returns>The Element Mapping Table</returns>
        private DataTable ReadElementMapping()
        {
            DataTable elementMappingTable;
            var assembly = Assembly.GetExecutingAssembly();
            Stream elementMappingStream = assembly.GetManifestResourceStream(this.elementMappingFileFullName);
            DataSet elementMappingDs = new DataSet();
            elementMappingDs.ReadXml(elementMappingStream);
            elementMappingTable = elementMappingDs.Tables["ElementMappingEntry"];
            return elementMappingTable;
        }
        #endregion
        
        #region Perform a Matching Test to ensure no Schema elements are missing from the Element Mapping File
        /// <summary>
        /// Perform a Matching Test to ensure no Schema elements are missing from the Element Mapping File
        /// </summary>
        /// <param name="elementMappingTable">The Element Mapping Table</param>
        private void PerformMatchingTest(DataTable elementMappingTable)
        {
            this.mainWindowObject.StatusLabel.Content = "Performing a Matching Test to ensure no Schema elements are missing from the Element Mapping File...";
            DoEvents();

            recordType matchingTest = new recordType();
            List<string> elementsNotInMappingFile = new List<string>();

            Type matchingType = matchingTest.GetType();
            PropertyInfo[] matchingProps = matchingType.GetProperties();
            foreach (PropertyInfo p in matchingProps)
            {
                string attributeName = string.Empty;
                if (p.GetCustomAttribute(typeof(XmlElementAttribute)) != null)
                {
                    XmlElementAttribute attr = (XmlElementAttribute)p.GetCustomAttribute(typeof(XmlElementAttribute));
                    attributeName = attr.ElementName;
                }
                if (attributeName != string.Empty)
                {
                    if (this.GetDataColumnName(attributeName, elementMappingTable) == null)
                    {
                        if (!p.Name.Contains("Specified"))
                        {
                            elementsNotInMappingFile.Add(attributeName);
                        }
                    }
                }
                else
                {
                    if (this.GetDataColumnName(p.Name, elementMappingTable) == null)
                    {
                        if (!p.Name.Contains("Specified"))
                        {
                            elementsNotInMappingFile.Add(p.Name);
                        }
                    }
                }
            }

            if (elementsNotInMappingFile.Count > 0)
            {
                StringBuilder sb = new StringBuilder();
                foreach (string element in elementsNotInMappingFile)
                {
                    sb.Append(element + Environment.NewLine);
                }

                MessageBox.Show("The following schema elements are not in the Mapping File and will be ignored:" + Environment.NewLine + Environment.NewLine + sb.ToString());
            }
        }
        #endregion
        
        #region Get the Results
        /// <summary>
        /// Get a list of results to be appended to the dataset object
        /// </summary>
        /// <param name="rows">The Data Rows from the supplied Data File </param>
        /// <param name="elementMappingTable">The Element Mapping Table</param>
        /// <returns>A list of XML Element result objects</returns>
        private List<recordType> GetResults(List<DataRow> rows, DataTable elementMappingTable)
        {
            this.mainWindowObject.StatusLabel.Content = "Generating XML ...";
            DoEvents();

            this.mainWindowObject.pbrDecoding.Value = 0;
            this.mainWindowObject.pbrDecoding.Minimum = 0;
            this.mainWindowObject.pbrDecoding.Maximum = rows.Count;

            try
            {
                List<recordType> results = new List<recordType>();

                foreach (DataRow row in rows)
                {
                    this.mainWindowObject.pbrDecoding.Value = rows.IndexOf(row) + 1;
                    DoEvents();

                    recordType resultObject = new recordType();

                    Type t = resultObject.GetType();
                    PropertyInfo[] props = t.GetProperties();
                    foreach (PropertyInfo p in props)
                    {
                        this.CreateXMLElement(p.Name, row, resultObject, elementMappingTable);
                    }

                    results.Add(resultObject);
                }

                return results;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error getting results: " + ex.Message);
                return null;
            }
        }
        #endregion

        #region Create the XML Element
        /// <summary>
        /// Create the XML Element
        /// </summary>
        /// <param name="resultElementName">The Name of the XML Element to be created</param>
        /// <param name="row">The DataRow from the DBF file</param>
        /// <param name="resultObject">The result object being added</param>
        /// <param name="elementMappingTable">The DataTable containing the Element Mapping</param>
        private void CreateXMLElement(string resultElementName, DataRow row, recordType resultObject, DataTable elementMappingTable)
        {
            // Get the DBF Column name from the codes.xml file
            string dataColumnName = this.GetDataColumnName(resultElementName, elementMappingTable);

            // Ensure that such a DBF Column exists 
            if (dataColumnName != null && row.Table.Columns.Contains(dataColumnName))
            {
                // Get the content of the DBF column
                string givenString = row[dataColumnName].ToString().Trim();

                // If there is conent, create the XML element, otherwise skip
                if (givenString != string.Empty)
                {
                    PropertyInfo prop = resultObject.GetType().GetProperty(resultElementName, BindingFlags.Public | BindingFlags.Instance);
                    if (prop != null && prop.CanWrite)
                    {
                        if (prop.PropertyType == typeof(string))
                        {
                            prop.SetValue(resultObject, givenString, null);
                        }

                        if (prop.PropertyType == typeof(decimal))
                        {
                            decimal decimalValue;
                            decimal.TryParse(givenString, out decimalValue);
                            prop.SetValue(resultObject, decimalValue, null);
                            PropertyInfo propSpecified = resultObject.GetType().GetProperty(resultElementName + "Specified", BindingFlags.Public | BindingFlags.Instance);
                            if (propSpecified != null && propSpecified.CanWrite && decimalValue != 0)
                            {
                                propSpecified.SetValue(resultObject, true, null);
                            }
                        }

                        if (prop.PropertyType == typeof(SSDCompoundType))
                        {
                            if (givenString.IndexOf("=") > -1)
                            {
                                // If there is a $ separator, there are multiple attributes
                                string[] attributeArray;
                                if (givenString.IndexOf("$") > -1)
                                {
                                    attributeArray = givenString.Split("$".ToCharArray());
                                }
                                else
                                {
                                    attributeArray = new string[1];
                                    attributeArray[0] = givenString;
                                }

                                List<SSDCompoundTypeValue> values = new List<SSDCompoundTypeValue>();
                                for (int i = 0; i < attributeArray.Length; i++)
                                {                                  
                                    SSDCompoundTypeValue ctv = new SSDCompoundTypeValue();
                                    ctv.name = attributeArray[i].Split("=".ToCharArray())[0];
                                    ctv.Value = attributeArray[i].Split("=".ToCharArray())[1];
                                    values.Add(ctv);  
                                }

                                SSDCompoundType newSSDCompoundType = new SSDCompoundType();
                                newSSDCompoundType.value = values.ToArray();
                                prop.SetValue(resultObject, newSSDCompoundType, null);
                            }
                            else
                            {
                                SSDCompoundType newSSDCompoundType = new SSDCompoundType();
                                newSSDCompoundType.Text = new string[1];
                                newSSDCompoundType.Text[0] = givenString;
                                prop.SetValue(resultObject, newSSDCompoundType, null);
                            }
                        }

                        if (prop.PropertyType == typeof(SSDRepeatableType))
                        {
                            SSDRepeatableType newSSDRepeatableType = new SSDRepeatableType();
                            //// List<String> rvalues = new List<String>();
                            //// rvalues.Add(givenString);
                            //// newSSDRepeatableType.value = rvalues.ToArray();
                            newSSDRepeatableType.Text = new string[1];
                            newSSDRepeatableType.Text[0] = givenString;
                            prop.SetValue(resultObject, newSSDRepeatableType, null);
                        }
                    }
                }
            }
        }
        #endregion

        #region Serialize the Message
        /// <summary>
        /// Serialize the Message
        /// </summary>
        /// <param name="messageObject">The Message to serialize</param>
        /// <param name="outputXLMFileEncoding">The Encoding of the XML file</param>
        /// <returns>A Boolean indicating whether the message has been Serialized</returns>
        private bool Serialize(message messageObject, System.Text.Encoding outputXLMFileEncoding)
        {
            this.mainWindowObject.StatusLabel.Content = "Writing XML ...";
            DoEvents();

            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(message));

                using (FileStream stream = new FileStream(this.outputXMLFileName, FileMode.Create))
                {
                    XmlTextWriter writer = new XmlTextWriter(stream, outputXLMFileEncoding);
                    writer.Formatting = Formatting.Indented;
                    writer.Indentation = 4;

                    serializer.Serialize(writer, messageObject);
                    writer.Close();
                }

                this.mainWindowObject.StatusLabel.Content = "XML generated. Now validating ...";
                DoEvents();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("XML Serialization error " + ex.Message);
                return false;
            }
        }
        #endregion
        
        #region Validate the produced XML file against the XSD
        /// <summary>
        /// Validate the produced XML file against the XSD
        /// </summary>
        private void Validate()
        {
            this.mainWindowObject.StatusLabel.Content = "Validating XML ...";
            DoEvents();

            if (File.Exists("errors.txt"))
            {
                File.Delete("errors.txt");
            }
                
            XmlReaderSettings settings = new XmlReaderSettings();
            settings.ValidationType = ValidationType.Schema;
            settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessInlineSchema;
            settings.ValidationFlags |= XmlSchemaValidationFlags.ProcessSchemaLocation;
            settings.ValidationFlags |= XmlSchemaValidationFlags.ReportValidationWarnings;
            settings.ValidationEventHandler += new ValidationEventHandler(this.ValidationCallBack);

            var assembly = Assembly.GetExecutingAssembly();
            Stream stream = assembly.GetManifestResourceStream(this.xsdFileFullName);

            XmlSchema schema;
            using (StreamReader sr = new StreamReader(stream))
            {
                var fs = sr.BaseStream;
                schema = XmlSchema.Read(fs, this.ValidationCallBack);
            }

            settings.Schemas.Add(schema);

            XmlReader reader = XmlReader.Create(this.outputXMLFileName, settings);
            try
            {
                while (reader.Read())
                {
                }

                if (this.validationErrorCount == 0)
                {
                    this.mainWindowObject.StatusLabel.Content = "XML generated and validated against the schema.";
                }
                else
                {
                    this.mainWindowObject.StatusLabel.Content = "The XML generated FAILED validation against the schema. Please see the errors file";
                }
            }
            catch (XmlException xex)
            {
                MessageBox.Show(xex.Message);
            }

            reader.Close();
        }

        /// <summary>
        /// Validation CallBack
        /// </summary>
        /// <param name="sender">The Sender Object</param>
        /// <param name="e">The Event Arguments</param>
        private void ValidationCallBack(object sender, ValidationEventArgs e)
        {
            string errorMessageFileName = "errors.txt";           
            File.AppendAllText(errorMessageFileName, "Failed XSD Validation - " + e.Message + Environment.NewLine);
            this.validationErrorCount++;
        }
        #endregion

        #region Get the Data Column name for the corresponding XML element name given
        /// <summary>
        /// Returns the Data Column name for the corresponding XML element name given
        /// </summary>
        /// <param name = "xmlElementName" > The XML Element Name</param>
        /// <param name = "elementMappingTable" > The XML Element Mapping Table</param>
        /// <returns>The Data Column Name</returns>
        private string GetDataColumnName(string xmlElementName, DataTable elementMappingTable)
        {
            var code = (from c in elementMappingTable.AsEnumerable()
                        where c.Field<string>("XMLElementName") == xmlElementName
                        select c.Field<string>("DataColumnName")).SingleOrDefault();
            return code;
        }
        #endregion

        #region Populate DataTable From Delimited File
        /// <summary>
        /// Populate DataTable From Delimited File
        /// </summary>
        /// <param name="filePath">The File Path of the delimited file</param>
        /// <returns>A DataTable Object from the Delimited data file</returns>
        private DataTable PopulateDataTableFromDelimitedFile(string filePath)
        {
            DataTable dataTableCSVdata = null;
            TextFieldParser textFieldParser = new TextFieldParser(filePath);

            try
            {
                switch (this.delimiterType)
                {
                    case DelimiterType.CommaWithQuotes:
                        textFieldParser.SetDelimiters(new string[] { "," });
                        textFieldParser.HasFieldsEnclosedInQuotes = true;
                        break;
                    case DelimiterType.Pipe:
                        textFieldParser.SetDelimiters(new string[] { "|" });
                        break;
                    case DelimiterType.Tab:
                        textFieldParser.SetDelimiters(new string[] { "\t" });
                        break;
                    case DelimiterType.Semicolon:
                        textFieldParser.SetDelimiters(new string[] { ";" });
                        break;
                }

                textFieldParser.TextFieldType = FieldType.Delimited;

                string[] columnHeaders = textFieldParser.ReadFields();

                DataTable dt = new DataTable();

                foreach (string columnHeader in columnHeaders)
                {
                    DataColumn dataColumn = new DataColumn(columnHeader);
                    dataColumn.AllowDBNull = true;
                    dt.Columns.Add(dataColumn);
                }

                // Iterate through the columns of each row
                int lineNumber = 0;
                while (!textFieldParser.EndOfData)
                {
                    lineNumber++;
                    string[] columns = textFieldParser.ReadFields();
                    for (int i = 0; i < columns.Length; i++)
                    {
                        if (columns[i] == string.Empty)
                        {
                            columns[i] = null;
                        }
                    }

                    if (columns.Length != columnHeaders.Length)
                    {
                        if (MessageBox.Show(
                            "Problem with line " + lineNumber.ToString() + "."
                            + Environment.NewLine
                            + Environment.NewLine
                            + "It possibly contains extra delimiters or it has fields enclosed in Quotes."
                            + Environment.NewLine
                            + Environment.NewLine
                            + "If you would like to just skip this line and proceed, press OK. "
                            + Environment.NewLine
                            + "If you would like to stop the XML generation, press Cancel",
                                "Problems with line " + lineNumber.ToString() + ".",
                                MessageBoxButton.OKCancel,
                                MessageBoxImage.Warning) == MessageBoxResult.Cancel)
                        {
                            break;
                        }                         
                    }
                    else
                    {
                        dt.Rows.Add(columns);
                    }
                }

                dataTableCSVdata = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
                this.mainWindowObject.StatusLabel.Content = string.Empty;
                this.mainWindowObject.pbrDecoding.Value = 0;
            }

            return dataTableCSVdata;
        }
        #endregion
    }
}
