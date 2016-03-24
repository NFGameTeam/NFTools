using System;
using System.Text;
using System.Windows.Forms;
using CCWin;
using CCWin.SkinControl;
using System.Xml;
using System.IO;
using System.Threading.Tasks;
using NPOI.XSSF.UserModel;

namespace FileProcess
{
    public partial class frmMain : Skin_Metro
    {
        int nCipher = 1;
        string strCipherCfg = "conf";

        string strExecutePath = "NFDataCfg/";
        string strToolBasePath = "../";
        string strRelativePath = "../../";
        string strExcelStructPath = "Excel_Struct/";
        string strXMLStructPath = "Struct/Class/";

        string strExcelIniPath = "Excel_Ini/";
        string strXMLIniPath = "Ini/NPC/";

        string strLogicClassFile = "";
        string strMySQLFile = "../mysql/NFrame.sql";
        string strMySQLClassFile = "../mysql/NFClass.sql";

        string strProtoFile = "../proto/NFRecordDefine.proto";

        String strHPPFile = "../proto/NFProtocolDefine.hpp";
        String strJavaFile = "../proto/NFProtocolDefine.java";
        String strCSFile = "../proto/NFProtocolDefine.cs";

        StreamWriter mysqlWriter;
        StreamWriter mysqlClassWriter;
        StreamWriter protoWriter;
        StreamWriter hppWriter;
        StreamWriter javaWriter;
        StreamWriter csWriter;

        String strHppIObjectInfo;
        String strJavaIObjectInfo;
        String strCSIObjectInfo;

        Int32 nRecordStart = 11;

        public frmMain()
        {
            InitializeComponent();
            //多线程访问控件需要加上这个
            CheckForIllegalCrossThreadCalls = false;
            if (File.Exists(strCipherCfg))
            {
                StreamReader confReader = new StreamReader(strCipherCfg);
                string strData = confReader.ReadToEnd();
                Int32.TryParse(strData, out nCipher);
            }

            if (nCipher > 0)
            {
                strLogicClassFile = "../Struct/LogicClass.NF";
            }
            else
            {
                strLogicClassFile = "../Struct/LogicClass.xml";
            }

        }

        private void OnCreateMysqlFile(object sender, EventArgs e)
        {
            this.skinButton1.Text = "Wait for generating MySQL file...";
            this.skinButton1.Enabled = false;

            CreateMySQLFile();
            CreateMySQLClassFile();

            if (!LoadLogicClass(strLogicClassFile))
            {
                MessageBox.Show("Generate MySQL file failed, please check files!");
                this.skinButton1.Text = "Generate MySQL file failed";
            }
            else
            {
                this.skinButton1.Text = "Generate MySQL file successful";
            }

            this.skinButton1.Enabled = true;

            mysqlWriter.Close();
            mysqlClassWriter.Close();
        }

        void CreateMySQLFile()
        {
            if (File.Exists(strMySQLFile))
            {
                File.Delete(strMySQLFile);
            }

            mysqlWriter = new StreamWriter(strMySQLFile);
        }

        void CreateMySQLClassFile()
        {
            if (File.Exists(strMySQLClassFile))
            {
                File.Delete(strMySQLClassFile);
            }

            mysqlClassWriter = new StreamWriter(strMySQLClassFile);
        }

        void CreateProtoFile()
        {
            if (File.Exists(strProtoFile))
            {
                File.Delete(strProtoFile);
            }

            protoWriter = new StreamWriter(strProtoFile);

            // 先把头写进去
            protoWriter.WriteLine("package NFrame;\n");
        }

        void CreateNameFile()
        {
            String strHPPHead = "// -------------------------------------------------------------------------\n"
            + "//    @FileName         :    NFProtocolDefine.hpp\n"
            + "//    @Author           :    NFrame Studio\n"
            + "//    @Date             :    " + DateTime.Now.ToString("yyyy/MM/dd") + "\n"
            + "//    @Module           :    NFProtocolDefine\n"
            + "// -------------------------------------------------------------------------\n\n"
            + "#ifndef NF_PR_NAME_HPP\n"
            + "#define NF_PR_NAME_HPP\n\n"
            + "#include <string>\n"
            + "namespace NFrame\n{\n";

            /////////////////////////////////////////////////////
            if (File.Exists(strHPPFile))
            {
                File.Delete(strHPPFile);
            }

            hppWriter = new StreamWriter(strHPPFile);
            hppWriter.WriteLine(strHPPHead);
            /////////////////////////////////////////////////////
            if (File.Exists(strJavaFile))
            {
                File.Delete(strJavaFile);
            }

            String strJavaHead = "// -------------------------------------------------------------------------\n"
            + "//    @FileName         :    NFProtocolDefine.java\n"
            + "//    @Author           :    NFrame Studio\n"
            + "//    @Date             :    " + DateTime.Now.ToString("yyyy/MM/dd") + "\n"
            + "//    @Module           :    NFProtocolDefine\n"
            + "// -------------------------------------------------------------------------\n\n"
            + "package nframe;\n";

            javaWriter = new StreamWriter(strJavaFile);
            javaWriter.WriteLine(strJavaHead);
            /////////////////////////////////////////////////////
            if (File.Exists(strCSFile))
            {
                File.Delete(strCSFile);
            }

            String strCSHead = "// -------------------------------------------------------------------------\n"
            + "//    @FileName         :    NFProtocolDefine.cs\n"
            + "//    @Author           :    NFrame Studio\n"
            + "//    @Date             :    " + DateTime.Now.ToString("yyyy/MM/dd") + "\n"
            + "//    @Module           :    NFProtocolDefine\n"
            + "// -------------------------------------------------------------------------\n\n"
            + "using System;\n"
            //+ "using System.Collections.Concurrent;\n"
            + "using System.Collections.Generic;\n"
            + "using System.Linq;\n"
            + "using System.Text;\n"
            + "using System.Threading;\n"
            //+ "using System.Threading.Tasks;\n\n"
            + "namespace NFrame\n{\n";

            csWriter = new StreamWriter(strCSFile);
            csWriter.WriteLine(strCSHead);
            /////////////////////////////////////////////////////
        }

        byte[] GetDataFromNFFile(string strFile)
        {
            StreamReader xReader = new StreamReader(strFile);
            string strData = xReader.ReadToEnd();
            xReader.Close();

            if (nCipher > 0)
            {
                return Convert.FromBase64String(strData);
            }
            else
            {
                return Encoding.UTF8.GetBytes(strData);
            }
        }

        bool LoadLogicClass(string strFile)
        {
            skinProgressBar1.Visible = false;
            skinProgressBar2.Value = 0;
            ////////////////////////////////////////////////////
            byte[] data = GetDataFromNFFile(strFile);
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(Encoding.UTF8.GetString(data));

            XmlElement root = doc.DocumentElement;
            XmlNode classNode = root.SelectSingleNode("Class");

            XmlElement classElement = (XmlElement)classNode;
            string strIObjectPath = classElement.GetAttribute("Path");

            XmlNodeList classNodeList = classNode.ChildNodes;
            skinProgressBar2.Maximum = classNodeList.Count;
            foreach (XmlNode node in classNodeList)
            {
                XmlElement nodeElement = (XmlElement)node;
                string strID = nodeElement.GetAttribute("Id");
                mysqlClassWriter.WriteLine("CREATE TABLE `" + strID + "` (");
                mysqlClassWriter.WriteLine("\t`ID` varchar(128) NOT NULL,");
                mysqlClassWriter.WriteLine("\tPRIMARY KEY (`ID`)");
                mysqlClassWriter.WriteLine(") ENGINE=InnoDB DEFAULT CHARSET=utf8;");

                // 读取父节点内容
                if (!LoadClass(strRelativePath + strIObjectPath, strID))
                {
                    return false;
                }

                // 读取自己节点内容
                string strPath = nodeElement.GetAttribute("Path");
                if (!LoadClass(strRelativePath + strPath, strID))
                {
                    return false;
                }

                mysqlClassWriter.WriteLine();

                skinProgressBar2.Value += 1;
            }

            return true;
        }

        bool LoadClass(string strFile, string strTable)
        {
            byte[] data = GetDataFromNFFile(strFile);
            XmlDocument xDoc = new XmlDocument();
            xDoc.LoadXml(Encoding.UTF8.GetString(data));

            XmlElement root = xDoc.DocumentElement;
            XmlNode xPropertysNode = root.SelectSingleNode("Propertys");
            XmlNodeList xPropertyNodeList = xPropertysNode.ChildNodes;
            foreach (XmlNode node in xPropertyNodeList)
            {
                XmlElement nodeElement = (XmlElement)node;
                if (nodeElement.GetAttribute("Save") != "1")
                {
                    continue;
                }

                string strID = nodeElement.GetAttribute("Id");
                string strDesc = nodeElement.GetAttribute("Desc");
                string strType = nodeElement.GetAttribute("Type");

                switch (strType)
                {
                    case "string":
                        mysqlWriter.WriteLine("ALTER TABLE `" + strTable + "` ADD `" + strID + "` varchar(128) DEFAULT '' COMMENT '" + strDesc + "';");
                        break;
                    case "int":
                        mysqlWriter.WriteLine("ALTER TABLE `" + strTable + "` ADD `" + strID + "` bigint(11) DEFAULT '0' COMMENT '" + strDesc + "';");
                        break;
                    case "object":
                        mysqlWriter.WriteLine("ALTER TABLE `" + strTable + "` ADD `" + strID + "` varchar(128) DEFAULT '' COMMENT '" + strDesc + "';");
                        break;
                    case "float":
                        mysqlWriter.WriteLine("ALTER TABLE `" + strTable + "` ADD `" + strID + "` float(11,3) DEFAULT '0' COMMENT '" + strDesc + "';");
                        break;
                    default:
                        mysqlWriter.WriteLine("ALTER TABLE `" + strTable + "` ADD `" + strID + "` varchar(128) DEFAULT '' COMMENT '" + strDesc + "';");
                        break;
                }
            }

            XmlNode xRecordsNode = root.SelectSingleNode("Records");
            XmlNodeList xRecordNodeList = xRecordsNode.ChildNodes;
            foreach (XmlNode node in xRecordNodeList)
            {
                XmlElement nodeElement = (XmlElement)node;
                if (nodeElement.GetAttribute("Save") != "1")
                {
                    continue;
                }

                string strID = nodeElement.GetAttribute("Id");
                string strDesc = nodeElement.GetAttribute("Desc");
                mysqlWriter.WriteLine("ALTER TABLE `" + strTable + "` ADD `" + strID + "` BLOB COMMENT '" + strDesc + "';");
            }

            return true;
        }

        /////////////////////////////////////////////////////////////////////////

        private void OnCreateXMLFile(object sender, EventArgs e)
        {
            this.skinButton2.Text = "Wait for generating NF files...";
            this.skinButton2.Enabled = false;

            //////////////////////////////////////////////////////////////////////
            Task xStructTask = new Task(CreateStructThreadFunc);
            xStructTask.Start();
            xStructTask.ContinueWith(CreateIniThreadFunc);
        }

        void CreateXMLCallBack()
        {
            this.skinButton2.Text = "Generate NF files successful";
            this.skinButton2.Enabled = true;

            this.skinButton1.Enabled = true;

            // proto文件结束
            protoWriter.Close();

            // name文件关闭
            hppWriter.Close();
            javaWriter.Close();
            csWriter.Close();

            strHppIObjectInfo = "";
            strJavaIObjectInfo = "";
            strCSIObjectInfo = "";
        }

        void CreateStructThreadFunc()
        {
            try
            {
                skinProgressBar1.Visible = true;

                // 生成proto文件
                CreateProtoFile();

                // 升成Property和Record名字hpp头文件
                CreateNameFile();

                XmlDocument xmlDoc = new XmlDocument();
                XmlDeclaration xmlDecl;
                xmlDecl = xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
                xmlDoc.AppendChild(xmlDecl);

                XmlElement root = xmlDoc.CreateElement("", "XML", "");
                xmlDoc.AppendChild(root);

                XmlElement classElement = xmlDoc.CreateElement("", "Class", "");
                root.AppendChild(classElement);
                //////////////////////////////////////////////////////////////////////
                classElement.SetAttribute("Id", "IObject");
                classElement.SetAttribute("Type", "TYPE_IOBJECT");
                if (nCipher > 0)
                {
                    classElement.SetAttribute("Path", strExecutePath + "Struct/Class/IObject.NF");
                    classElement.SetAttribute("InstancePath", strExecutePath + "Ini/NPC/IObject.NF");
                }
                else
                {
                    classElement.SetAttribute("Path", strExecutePath + "Struct/Class/IObject.xml");
                    classElement.SetAttribute("InstancePath", strExecutePath + "Ini/NPC/IObject.xml");
                }

                classElement.SetAttribute("Public", "0");
                classElement.SetAttribute("Desc", "IObject");
                //////////////////////////////////////////////////////////////////////
                // 提前把IObject跑一边
                CreateStructXML("../Excel_Struct/IObject.xlsx", "IObject");

                // 遍历Struct文件夹下的excel文件
                String[] xStructXLSXList = Directory.GetFiles(strToolBasePath + strExcelStructPath, "*", SearchOption.AllDirectories);
                skinProgressBar1.Maximum = xStructXLSXList.Length - 1; // 去掉IObject
                skinProgressBar1.Value = 0;
                foreach (string file in xStructXLSXList)
                {
                    int nLastPoint = file.LastIndexOf(".") + 1;
                    int nLastSlash = file.LastIndexOf("/") + 1;
                    string strFileName = file.Substring(nLastSlash, nLastPoint - nLastSlash - 1);
                    string strFileExt = file.Substring(nLastPoint, file.Length - nLastPoint);

                    int nTempExcelPos = file.IndexOf("$");
                    if (nTempExcelPos >= 0)
                    {
                        // 打开excel之后生成的临时文件，略过
                        skinProgressBar1.Maximum -= 1;
                        continue;
                    }

                    // 不是excel文件，默认跳过
                    if (strFileExt != "xls" && strFileExt != "xlsx")
                    {
                        skinProgressBar1.Maximum -= 1;
                        continue;
                    }

                    // 是IObject.xlsx跳过
                    if (strFileName == "IObject")
                    {
                        continue;
                    }

                    // 单个excel文件转为xml
                    if (!CreateStructXML(file, strFileName))
                    {
                        MessageBox.Show("Create " + file + " failed!");
                        return;
                    }

                    // LogicClass.xml中的索引配置
                    XmlElement subClassElement = xmlDoc.CreateElement("", "Class", "");
                    classElement.AppendChild(subClassElement);

                    subClassElement.SetAttribute("Id", strFileName);
                    subClassElement.SetAttribute("Type", "TYPE_" + strFileName.ToUpper());
                    if (nCipher > 0)
                    {
                        subClassElement.SetAttribute("Path", strExecutePath + strXMLStructPath + strFileName + ".NF");
                        subClassElement.SetAttribute("InstancePath", strExecutePath + strXMLIniPath + strFileName + ".NF");
                    }
                    else
                    {
                        subClassElement.SetAttribute("Path", strExecutePath + strXMLStructPath + strFileName + ".xml");
                        subClassElement.SetAttribute("InstancePath", strExecutePath + strXMLIniPath + strFileName + ".xml");
                    }


                    subClassElement.SetAttribute("Public", "0");
                    subClassElement.SetAttribute("Desc", strFileName);

                    // ProgressBar
                    skinProgressBar1.Value += 1;
                }

                xmlDoc.Save(strLogicClassFile);
                ProcessEncryptFile(strLogicClassFile);

                // name文件结束
                hppWriter.WriteLine("} // !@NFrame\n\n#endif // !NF_PR_NAME_HPP");
                // java 不需要namespace，所以不需要}
                csWriter.WriteLine("}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Some exception is found, error:{0}", ex.Message);
            }
        }

        void CreateIniThreadFunc(Task arg)
        {
            try
            {
                // 遍历Ini文件夹下的excel文件
                String[] xIniXLSXList = Directory.GetFiles(strToolBasePath + strExcelIniPath, "*", SearchOption.AllDirectories);
                skinProgressBar2.Maximum = xIniXLSXList.Length;
                skinProgressBar2.Value = 0;
                foreach (string file in xIniXLSXList)
                {
                    int nLastPoint = file.LastIndexOf(".") + 1;
                    int nLastSlash = file.LastIndexOf("/") + 1;
                    string strFileName = file.Substring(nLastSlash, nLastPoint - nLastSlash - 1);
                    string strFileExt = file.Substring(nLastPoint, file.Length - nLastPoint);

                    int nTempExcelPos = file.IndexOf("$");
                    if (nTempExcelPos >= 0)
                    {
                        // 打开excel之后生成的临时文件，略过
                        skinProgressBar2.Maximum -= 1;
                        continue;
                    }

                    // 不是excel文件，默认跳过
                    if (strFileExt != "xls" && strFileExt != "xlsx")
                    {
                        skinProgressBar2.Maximum -= 1;
                        continue;
                    }

                    if (!CreateIniXML(file))
                    {
                        MessageBox.Show("Create " + file + " failed!");
                        return;
                    }

                    skinProgressBar2.Value += 1;
                }

                Task xIniTask = new Task(CreateXMLCallBack);
                xIniTask.Start();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Some exception is found, error:{0}", ex.Message);
            }
        }

        bool CreateStructXML(string strFile, string strFileName)
        {
            try
            {
                string strCurPath = Directory.GetCurrentDirectory();

                // 打开excel
                FileStream file = new FileStream(strCurPath + "/" + strFile, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(file);

                // PropertyName
                // cpp
                String strHPPPropertyInfo = "";
                String strHppRecordInfo = "";
                String strHppEnumInfo = "";

                strHPPPropertyInfo += "class " + strFileName + "\n{\npublic:\n";
                strHPPPropertyInfo += "\t//Class name\n\tstatic const std::string& ThisName(){ static std::string x" + strFileName + " = \"" + strFileName + "\";" + " return x" + strFileName + "; }\n" + "\t// IObject\n" + strHppIObjectInfo + "\t// Property\n";

                // java
                String strJavaPropertyInfo = "";
                String strJavaRecordInfo = "";
                String strJavaEnumInfo = "";

                strJavaPropertyInfo += "public class " + strFileName + " {\n";
                strJavaPropertyInfo += "\t//Class name\n\tpublic static final String ThisName = \"" + strFileName + "\";\n\t// IObject\n" + strJavaIObjectInfo + "\t// Property\n";

                // C#
                String strCSPropertyInfo = "";
                String strCSRecordInfo = "";
                String strCSEnumInfo = "";

                strCSPropertyInfo += "public class " + strFileName + "\n{\n";
                strCSPropertyInfo += "\t//Class name\n\tpublic static readonly string ThisName = \"" + strFileName + "\";\n\t// IObject\n" + strCSIObjectInfo + "\t// Property\n";

                // 开始创建xml
                XmlDocument structDoc = new XmlDocument();
                XmlDeclaration xmlDecl;
                xmlDecl = structDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
                structDoc.AppendChild(xmlDecl);

                // 写入XML root标签
                XmlElement root = structDoc.CreateElement("", "XML", "");
                structDoc.AppendChild(root);

                // 写入Propertys标签
                XmlElement propertyNodes = structDoc.CreateElement("", "Propertys", "");
                root.AppendChild(propertyNodes);

                // 写入Records标签
                XmlElement recordNodes = structDoc.CreateElement("", "Records", "");
                root.AppendChild(recordNodes);

                // 写入components的处理
                XmlElement componentNodes = structDoc.CreateElement("", "Components", "");
                root.AppendChild(componentNodes);

                // 读取excel中每一个sheet
                //foreach (Worksheet sheet in book.Worksheets)
                foreach (XSSFSheet sheet in wb)
                {
                    string strSheetName = sheet.SheetName;
                    int nRowCount = sheet.LastRowNum;
                    int nColCount = sheet.GetRow(0).LastCellNum;

                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                    string[] colNames = new string[nColCount];
                    rows.MoveNext();
                    for (int i = 0; i < nColCount; i++)
                    {
                        XSSFRow row = rows.Current as XSSFRow;
                        var ss = row.GetCell(i).StringCellValue;
                        colNames[i] = ss;
                    }

                    switch (strSheetName.ToLower())
                    {
                        case "property":
                            {
                                while (rows.MoveNext())
                                {
                                    XSSFRow row = rows.Current as XSSFRow;
                                    string testValue = row.GetCell(0).StringCellValue;
                                    if (testValue.IsNullOrEmpty())
                                    {
                                        continue;
                                    }

                                    var propertyNode = structDoc.CreateElement("", "Property", "");
                                    propertyNodes.AppendChild(propertyNode);

                                    string strType = "";
                                    for (Int32 col = 0; col < nColCount; ++col)
                                    {
                                        try
                                        {
                                            string name = colNames[col];

                                            string value = "";
                                            if (!row.GetCell(col).IsNull())
                                            {
                                                var valueCell = row.GetCell(col);
                                                if (valueCell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                                                {
                                                    value = row.GetCell(col).BooleanCellValue ? "1" : "0";
                                                }
                                                else if (valueCell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                                                {
                                                    value = row.GetCell(col).NumericCellValue.ToString();
                                                }
                                                else
                                                {

                                                    value = row.GetCell(col).StringCellValue;
                                                }

                                                if (name == "Type")
                                                {
                                                    strType = value;
                                                }
                                            }
                                            propertyNode.SetAttribute(name, value.ToString());
                                        }
                                        catch (Exception err)
                                        {
                                            MessageBox.Show(strFile + "\n" + "Sheet: " + strSheetName + "\n" + "Cell: " + row.RowNum.ToString() + ", " + col.ToString() + "\n" + "This cell is empty or error\n" + err.Message);
                                            return false;
                                        }
                                    }

                                    strHPPPropertyInfo += "\tstatic const std::string& " + testValue + "(){ static std::string x" + testValue + " = \"" + testValue + "\";" + " return x" + testValue + "; } // " + strType + "\n";
                                    strJavaPropertyInfo += "\tpublic static final String " + testValue + " = \"" + testValue + "\"; // " + strType + "\n";
                                    strCSPropertyInfo += "\tpublic static readonly String " + testValue + " = \"" + testValue + "\"; // " + strType + "\n";

                                    if (strFileName == "IObject")
                                    {
                                        strHppIObjectInfo += "\tstatic const std::string& " + testValue + "(){ static std::string x" + testValue + " = \"" + testValue + "\";" + " return x" + testValue + "; } // " + strType + "\n";
                                        strJavaIObjectInfo += "\tpublic static final String " + testValue + " = \"" + testValue + "\"; // " + strType + "\n";
                                        strCSIObjectInfo += "\tpublic static readonly String " + testValue + " = \"" + testValue + "\"; // " + strType + "\n";
                                    }
                                }

                                if (strFileName == "IObject")
                                {

                                }
                            }
                            break;
                        case "component":
                            {
                                while (rows.MoveNext())
                                {
                                    XSSFRow row = rows.Current as XSSFRow;
                                    string testValue = row.GetCell(0).StringCellValue;
                                    if (testValue.IsNullOrEmpty())
                                    {
                                        continue;
                                    }

                                    var componentNode = structDoc.CreateElement("", "Component", "");
                                    componentNodes.AppendChild(componentNode);

                                    string strType = "";
                                    for (Int32 col = 0; col < nColCount; ++col)
                                    {
                                        try
                                        {
                                            string name = colNames[col];

                                            string value = "";
                                            if (!row.GetCell(col).IsNull())
                                            {
                                                var valueCell = row.GetCell(col);
                                                if (valueCell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                                                {
                                                    value = row.GetCell(col).BooleanCellValue ? "1" : "0";
                                                }
                                                else if (valueCell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                                                {
                                                    value = row.GetCell(col).NumericCellValue.ToString();
                                                }
                                                else
                                                {

                                                    value = row.GetCell(col).StringCellValue;
                                                }

                                                if (name == "Type")
                                                {
                                                    strType = value;
                                                }
                                            }
                                            componentNode.SetAttribute(name, value.ToString());
                                        }
                                        catch (Exception err)
                                        {
                                            MessageBox.Show(strFile + "\n" + "Sheet: " + strSheetName + "\n" + "Cell: " + row.RowNum.ToString() + ", " + col.ToString() + "\n" + "This cell is empty or error\n" + err.Message);
                                            return false;
                                        }
                                    }
                                }
                            }
                            break;
                        default:
                            {
                                var recordNode = structDoc.CreateElement("", "Record", "");
                                recordNodes.AppendChild(recordNode);

                                string strRecordName = "";

                                // 从11个开始就是Record中具体的Col了
                                Int32 nSetColNum = 0;
                                rows.MoveNext();
                                XSSFRow valueRow = rows.Current as XSSFRow;
                                for (Int32 col = 0; col < nRecordStart; ++col)
                                {
                                    try
                                    {
                                        string name = colNames[col];
                                        string value = "";
                                        if (!valueRow.GetCell(col).IsNull())
                                        {
                                            var valueCell = valueRow.GetCell(col);
                                            if (valueCell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                                            {
                                                value = valueRow.GetCell(col).BooleanCellValue ? "1" : "0";
                                            }
                                            else if (valueCell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                                            {
                                                value = valueRow.GetCell(col).NumericCellValue.ToString();
                                            }
                                            else
                                            {

                                                value = valueRow.GetCell(col).StringCellValue;
                                            }
                                            recordNode.SetAttribute(name, value);
                                            if (name == "Col")
                                            {
                                                nSetColNum = Int32.Parse(value);
                                            }

                                            if (name == "Id")
                                            {
                                                strRecordName = value;
                                            }
                                        }
                                    }
                                    catch (Exception err)
                                    {
                                        MessageBox.Show(strFile + "\n" + "Sheet: " + strSheetName + "\n" + "Cell: 1/2 " + col.ToString() + "\n" + "This cell is empty or error\n" + err.Message);
                                        return false;
                                    }
                                }

                                strHppRecordInfo += "\tstatic const std::string& R_" + strRecordName + "(){ static std::string x" + strRecordName + " = \"" + strRecordName + "\";" + " return x" + strRecordName + ";}\n";
                                strHppEnumInfo += "\n\tenum " + strRecordName + "\n\t{\n";

                                strJavaRecordInfo += "\tpublic static final String R_" + strRecordName + " = \"" + strRecordName + "\";\n";
                                strJavaEnumInfo += "\n\tpublic enum " + strRecordName + "\n\t{\n";

                                strCSRecordInfo += "\tpublic static readonly String R_" + strRecordName + " = \"" + strRecordName + "\";\n";
                                strCSEnumInfo += "\n\tpublic enum " + strRecordName + "\n\t{\n";

                                protoWriter.WriteLine("enum " + strRecordName);
                                protoWriter.WriteLine("{");

                                for (Int32 col = nRecordStart; col < nRecordStart + nSetColNum; ++col)
                                {
                                    try
                                    {
                                        string name = colNames[col];
                                        string value = "";
                                        if (!valueRow.GetCell(col).IsNull())
                                        {
                                            var valueCell = valueRow.GetCell(col);
                                            if (valueCell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                                            {
                                                value = valueRow.GetCell(col).BooleanCellValue ? "1" : "0";
                                            }
                                            else if (valueCell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                                            {
                                                value = valueRow.GetCell(col).NumericCellValue.ToString();
                                            }
                                            else
                                            {

                                                value = valueRow.GetCell(col).StringCellValue;
                                            }
                                            var colNode = structDoc.CreateElement("Col");
                                            recordNode.AppendChild(colNode);

                                            colNode.SetAttribute("Type", value);
                                            colNode.SetAttribute("Tag", name);

                                            protoWriter.WriteLine("\t" + name + "\t\t= " + (col - nRecordStart).ToString() + "; // " + name + " -- " + value);

                                            if (col == nRecordStart + nSetColNum - 1)
                                            {
                                                strHppEnumInfo += "\t\t" + strRecordName + "_" + name + "\t\t= " + (col - nRecordStart).ToString() + ", // " + name + " -- " + value;
                                                strJavaEnumInfo += "\t\t" + name + "\t\t= " + (col - nRecordStart).ToString() + ", // " + name + " -- " + value;
                                                strCSEnumInfo += "\t\t" + name + "\t\t= " + (col - nRecordStart).ToString() + ", // " + name + " -- " + value;
                                            }
                                            else
                                            {
                                                strHppEnumInfo += "\t\t" + strRecordName + "_" + name + "\t\t= " + (col - 11).ToString() + ", // " + name + " -- " + value + "\n";
                                                strJavaEnumInfo += "\t\t" + name + "\t\t= " + (col - nRecordStart).ToString() + ", // " + name + " -- " + value + "\n";
                                                strCSEnumInfo += "\t\t" + name + "\t\t= " + (col - nRecordStart).ToString() + ", // " + name + " -- " + value + "\n";
                                            }
                                        }
                                    }
                                    catch (Exception err)
                                    {
                                        MessageBox.Show(strFile + "\n" + "Sheet: " + strSheetName + "\n" + "Cell: 1/2 " + col.ToString() + "\n" + "This cell is empty or error\n" + err.Message);
                                        return false;
                                    }
                                }

                                protoWriter.WriteLine("}");
                                protoWriter.WriteLine();

                                strHppEnumInfo += "\n\t};\n";
                                strJavaEnumInfo += "\n\t};\n";
                                strCSEnumInfo += "\n\t};\n";
                            }
                            break;
                    }

                    // cpp
                    strHPPPropertyInfo += "\t// Record\n" + strHppRecordInfo + strHppEnumInfo + "\n};\n\n";
                    hppWriter.Write(strHPPPropertyInfo);

                    // java
                    strJavaPropertyInfo += "\t// Record\n" + strJavaRecordInfo + strJavaEnumInfo + "\n}\n\n";
                    javaWriter.Write(strJavaPropertyInfo);

                    // C#
                    strCSPropertyInfo += "\t// Record\n" + strCSRecordInfo + strCSEnumInfo + "\n}\n\n";
                    csWriter.Write(strCSPropertyInfo);

                    ////////////////////////////////////////////////////////////////////////////
                    // 保存文件
                    int nLastPoint = strFile.LastIndexOf(".") + 1;
                    int nLastSlash = strFile.LastIndexOf("/") + 1;
                    string strFileExt = strFile.Substring(nLastPoint, strFile.Length - nLastPoint);

                    string strXMLFile = strToolBasePath + strXMLStructPath + strFileName;
                    if (nCipher > 0)
                    {
                        strXMLFile += ".NF";
                    }
                    else
                    {
                        strXMLFile += ".xml";
                    }
                    structDoc.Save(strXMLFile);
                    ProcessEncryptFile(strXMLFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Some exception is found, File={0}, error:{1}", strFile, ex.Message);
                return false;
            }

            return true;
        }

        bool CreateIniXML(string strFile)
        {
            try
            {
                string strCurPath = Directory.GetCurrentDirectory();
                // 打开excel
                FileStream file = new FileStream(strCurPath + "/" + strFile, FileMode.Open, FileAccess.Read);
                XSSFWorkbook wb = new XSSFWorkbook(file);
                ////////////////////////////////////////////////////////////////////////////
                XmlDocument iniDoc = new XmlDocument();
                XmlDeclaration xmlDecl;
                xmlDecl = iniDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
                iniDoc.AppendChild(xmlDecl);

                XmlElement root = iniDoc.CreateElement("", "XML", "");
                iniDoc.AppendChild(root);

                // 读取excel中每一个sheet
                foreach (XSSFSheet sheet in wb)
                {
                    string strSheetName = sheet.SheetName;
                    int nRowCount = sheet.LastRowNum;
                    int nColCount = sheet.GetRow(0).LastCellNum;

                    System.Collections.IEnumerator rows = sheet.GetRowEnumerator();
                    string[] colNames = new string[nColCount];
                    rows.MoveNext();
                    bool isEmpty = false;
                    for (int i = 0; i < nColCount; i++)
                    {
                        try
                        {                            
                            if(((XSSFRow)rows.Current).GetCell(i).StringCellValue!="")
                            {
                                colNames[i] = ((XSSFRow)rows.Current).GetCell(i).StringCellValue;
                            }
                            else
                            {
                                isEmpty = true;
                                break;
                            }
                        }
                        catch(Exception err)
                        {
                            isEmpty = true;
                            break;
                        }
                    }
                    if(isEmpty)
                    {
                        break;
                    }                     
                    while (rows.MoveNext())
                    {
                        XSSFRow row = rows.Current as XSSFRow;
                        string testValue = "";
                        try
                        {
                            testValue = row.GetCell(0).StringCellValue;
                        }
                        catch(Exception err)
                        {
                            continue;
                        }
                        
                        if (testValue.IsNullOrEmpty())
                        {
                            continue;
                        }
                        XmlElement objectNode = iniDoc.CreateElement("", "Object", "");
                        root.AppendChild(objectNode);
                        for (Int32 col = 0; col < nColCount; ++col)
                        {
                            try
                            {
                                string name = colNames[col];

                                string value = "";
                                if (!row.GetCell(col).IsNull())
                                {
                                    var valueCell = row.GetCell(col);
                                    if (valueCell.CellType == NPOI.SS.UserModel.CellType.Boolean)
                                    {
                                        value = row.GetCell(col).BooleanCellValue ? "1" : "0";
                                    }
                                    else if (valueCell.CellType == NPOI.SS.UserModel.CellType.Numeric)
                                    {
                                        value = row.GetCell(col).NumericCellValue.ToString();
                                    }
                                    else
                                    {

                                        value = row.GetCell(col).StringCellValue;
                                    }

                                }
                                objectNode.SetAttribute(name, value.ToString());
                            }
                            catch (Exception err)
                            {
                                MessageBox.Show(strFile + "\n" + "Sheet: " + strSheetName + "\n" + "Cell: " + row.ToString() + ", " + col.ToString() + "\n" + "This cell is empty or error\n" + err.Message);
                                return false;
                            }
                        }
                    }
                }

                ////////////////////////////////////////////////////////////////////////////
                // 保存文件
                int nLastPoint = strFile.LastIndexOf(".") + 1;
                int nLastSlash = strFile.LastIndexOf("/") + 1;
                string strFileName = strFile.Substring(nLastSlash, nLastPoint - nLastSlash - 1);
                string strFileExt = strFile.Substring(nLastPoint, strFile.Length - nLastPoint);

                string strXMLFile = strToolBasePath + strXMLIniPath + strFileName;
                if (nCipher > 0)
                {
                    strXMLFile += ".NF";
                }
                else
                {
                    strXMLFile += ".xml";
                }

                //CloseFile(strXMLFile);
                iniDoc.Save(strXMLFile);

                ProcessEncryptFile(strXMLFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Some exception is found, File={0}, error:{1}", strFile, ex.Message));
                Console.WriteLine("Some exception is found, File={0}, error:{1}", strFile, ex.Message);
                return false;
            }

            return true;
        }

        void ProcessEncryptFile(string strFile)
        {
            if (nCipher <= 0)
            {
                return;
            }

            if (!File.Exists(strFile))
            {
                return;
            }

            StreamReader fileReader = new StreamReader(strFile);
            string strContent = fileReader.ReadToEnd();
            fileReader.Close();
            File.Delete(strFile);

            StreamWriter fileWriter = new StreamWriter(strFile);
            fileWriter.Write(Convert.ToBase64String(Encoding.UTF8.GetBytes(strContent)));
            fileWriter.Close();
        }
    }
}
