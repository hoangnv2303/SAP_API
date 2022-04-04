using SAP.Middleware.Connector;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Web;

namespace WebAPISAP.Common
{
    public class SAPHelper
    {
        private readonly string _SAPConnectionString;

        public SAPHelper()
        {
            _SAPConnectionString = ConfigurationManager.AppSettings["SAPConnectionString"];
        }

        public string ConfirmTO(DataVM entity, string confirm)
        {
            try
            {
                RfcConfigParameters rfcParam = Connector3SapRfcConnGroup(_SAPConnectionString);
                RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcParam);
                RfcRepository rfcRepository = rfcDest.Repository;
                IRfcFunction rfcFunction = rfcRepository.CreateFunction("L_TO_CONFIRM");
                rfcFunction.SetValue("I_LGNUM", "049");
                rfcFunction.SetValue("I_TANUM", entity.TO);
                rfcFunction.SetValue("I_QNAME", "RFCSHARE02");
                rfcFunction.SetValue("I_ENAME", "RFCSHARE02");
                rfcFunction.SetValue("I_COMMIT_WORK", "X");
                rfcFunction.SetValue("I_QUKNZ", confirm);

                IRfcTable table = rfcFunction.GetTable("T_LTAP_CONF");

                table.Insert();
                table.SetValue("TANUM", Convert.ToInt32(entity.TO));
                table.SetValue("TAPOS", entity.TO_ORDER);
                table.SetValue("ALTME", entity.UNIT);//unit of material code
                table.SetValue("NISTA", entity.QTY);
                //table.SetValue("NDIFA", "");//blank
                //table.SetValue("NLPLA", "");//blank

                rfcFunction.Invoke(rfcDest);
                return "Ok";
            }
            catch (RfcCommunicationException ex)
            {
                return "Error: " + ex.Message.ToString();
            }
            catch (RfcAbapException ex)
            {
                return "Error: " + ex.Message.ToString();
            }
            catch (RfcLogonException ex)
            {
                return "Error: " + ex.Message.ToString();
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message.ToString();
            }
        }

        public RfcConfigParameters Connector3SapRfcConnGroup(string connStr)
        {
            try
            {
                string[] arrConn = connStr.Split(' ');
                RfcConfigParameters rfcPar = new RfcConfigParameters();
                rfcPar.Clear();
                //rfcPar.Add(RfcConfigParameters.Name, DateTime.Now.ToString("yyyyMMddHHmmss"));
                rfcPar.Add(RfcConfigParameters.AppServerHost, arrConn[0].Split('=')[1]);
                rfcPar.Add(RfcConfigParameters.Name, arrConn[1].Split('=')[1]);
                rfcPar.Add(RfcConfigParameters.LogonGroup, arrConn[2].Split('=')[1]);
                rfcPar.Add(RfcConfigParameters.Client, arrConn[3].Split('=')[1]);
                rfcPar.Add(RfcConfigParameters.User, arrConn[4].Split('=')[1]);
                rfcPar.Add(RfcConfigParameters.Password, arrConn[5].Split('=')[1]);
                rfcPar.Add(RfcConfigParameters.Language, arrConn[6].Split('=')[1]);
                //rfcPar.Add(RfcConfigParameters.RepositoryConnectionIdleTimeout, "1");
                //rfcPar.Add(RfcConfigParameters.ConnectionIdleTimeout, "1");
                //rfcPar.Add(RfcConfigParameters.PeakConnectionsLimit, "1000");
                WriteLogs.Write("===", "Connect Ok");
                return rfcPar;
            }
            catch (RfcCommunicationException ex)
            {
                WriteLogs.Write("Connector3SapRfcConnGroup", ex);
                return null;
            }
            catch (RfcAbapException ex)
            {
                WriteLogs.Write("Connector3SapRfcConnGroup", ex);
                return null;
            }
            catch (RfcLogonException ex)
            {
                WriteLogs.Write("Connector3SapRfcConnGroup", ex);
                return null;
            }
            catch (Exception ex)
            {
                WriteLogs.Write("Connector3SapRfcConnGroup", ex);
                return null;
            }
        }
        public DataTable ConvertToDoNetTable(IRfcTable RFCTable)
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < RFCTable.ElementCount; i++)
            {
                RfcElementMetadata metadata = RFCTable.GetElementMetadata(i);
                dt.Columns.Add(metadata.Name);
            }
            foreach (IRfcStructure row in RFCTable)
            {
                DataRow dr = dt.NewRow();
                for (int i = 0; i < RFCTable.ElementCount; i++)
                {
                    RfcElementMetadata metadata = RFCTable.GetElementMetadata(i);
                    dr[i] = row.GetString(metadata.Name);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public string CreateTO(string fromLine, string toLine, string fromBin, int number, string part, string plant, int qty)
        {
            try
            {
                RfcConfigParameters rfcParam = Connector3SapRfcConnGroup(_SAPConnectionString);
                RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcParam);
                RfcRepository rfcRep = rfcDest.Repository;

                IRfcFunction rfcFun = rfcRep.CreateFunction("L_TO_CREATE_SINGLE");//SAP RFC

                // COMMON
                rfcFun.SetValue("I_LGNUM", "049");
                rfcFun.SetValue("I_BWLVS", number); // special for IQC
                rfcFun.SetValue("I_BESTQ", "Q"); // special for IQC

                // MATERIAL
                rfcFun.SetValue("I_MATNR", part); //part
                rfcFun.SetValue("I_ANFME", qty); //requested qty

                // LOCATION - BIN
                // Location: HHR -> IQC
                rfcFun.SetValue("I_VLTYP", fromLine); //line return 
                rfcFun.SetValue("I_NLTYP", toLine); //line destination
                // Bin:
                rfcFun.SetValue("I_VLPLA", fromBin); //requested qty
                //rfcFun.SetValue("I_BETYP", ""); // Auto empty bin


                //rfcFun.SetValue("I_BESTQ", "Q"); // IQC 
                //rfcFun.SetValue("I_LETYP", "BB"); // bin type BT

                rfcFun.SetValue("I_BETYP", "N");
                rfcFun.SetValue("I_BENUM", "IQC Inspection"); // Remark
                rfcFun.SetValue("I_WERKS", plant); //Plant

                rfcFun.Invoke(rfcDest);

                IRfcTable rtSource = rfcFun.GetTable("T_LTAP_VB");
                DataTable dtMaterialMaster = ConvertToDoNetTable(rtSource);
                if (dtMaterialMaster == null || dtMaterialMaster.Rows.Count > 0)
                {
                    string to = dtMaterialMaster.Rows[0]["TANUM"].ToString();
                    string unit = dtMaterialMaster.Rows[0]["ALTME"].ToString();
                    string desBin = dtMaterialMaster.Rows[0]["NLPLA"].ToString();
                    return to + ';' + unit + ';' + desBin;
                }
                else
                {
                    return "Error create TO, Please re-check!";
                }
            }
            catch (RfcCommunicationException ex)
            {
                return "Error: " + ex.Message.ToString();
            }
            catch (RfcAbapException ex)
            {
                return "Error: " + ex.Message.ToString();
            }
            catch (RfcLogonException ex)
            {
                return "Error: " + ex.Message.ToString();
            }
            catch (Exception ex)
            {
                return "Error: " + ex.Message.ToString();
            }
        }


        public DataTable DownloadBoom(string plant, string parentMaterial)
        {
            try
            {
                RfcConfigParameters rfcParam = Connector3SapRfcConnGroup(_SAPConnectionString);
                RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcParam);
                RfcRepository rfcRep = rfcDest.Repository;
                IRfcFunction rfcFun = rfcRep.CreateFunction("ZRFC_CS_BOM_EXPL_0003");//SAP RFC
                rfcFun.SetValue("WERKS", plant);
                rfcFun.SetValue("DATUV", DateTime.Now.ToString("yyyyMMdd"));
                rfcFun.SetValue("MEHRS", "X");
                rfcFun.SetValue("MDMPS", "");

                // Get table PLANT
                IRfcTable rtSourceOuput = null;
                IRfcTable rtSource1 = rfcFun.GetTable("MATNR_IN");
                rtSource1.Append();

                rtSource1.SetValue("MATNR", parentMaterial); // 16KIGLW1A13

                rfcFun.Invoke(rfcDest);
                rtSourceOuput = rfcFun.GetTable("BOM_INFO");

                DataTable dtBoom = ConvertToDoNetTable(rtSourceOuput);
                return dtBoom;
            }
            catch (RfcCommunicationException ex)
            {
                WriteLogs.Write("DownloadBoom", ex);
                return null;
            }
            catch (RfcAbapException ex)
            {
                WriteLogs.Write("DownloadBoom", ex);
                return null;
            }
            catch (RfcLogonException ex)
            {
                WriteLogs.Write("DownloadBoom", ex);
                return null;
            }
            catch (Exception ex)
            {
                WriteLogs.Write("DownloadBoom", ex);
                return null;
            }
        }

        public DataTable DownloadDeliverySale(string plant, DateTime firstDate, DateTime lastDate)
        {
            try
            {
                RfcConfigParameters rfcParam = Connector3SapRfcConnGroup(_SAPConnectionString);
                RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcParam);
                RfcRepository rfcRep = rfcDest.Repository;

                IRfcFunction rfcFun = rfcRep.CreateFunction("ZRFC_SD_ZPSD01");//SAP RFC
                rfcFun.SetValue("ZCCS0", "X");
                rfcFun.SetValue("P_HMD", "X");

                DataTable loTable = new DataTable();


                // Get table PLANT
                IRfcTable rtSource = rfcFun.GetTable("PLANT");
                // FLANT
                rtSource.Append();
                rtSource.SetValue("SIGN", "I");
                rtSource.SetValue("OPTION", "EQ");
                rtSource.SetValue("LOW", plant);

                // SALES_ORGANIZATION
                IRfcTable rtSourceDoc = rfcFun.GetTable("SALES_ORGANIZATION");
                rtSourceDoc.Append();
                rtSourceDoc.SetValue("SIGN", "I");
                rtSourceDoc.SetValue("OPTION", "EQ");
                rtSourceDoc.SetValue("LOW", "FI03");


                // Document date
                IRfcTable rtSourceDocDate = rfcFun.GetTable("DOCUMENT_DATE");

                rtSourceDocDate.Append();
                rtSourceDocDate.SetValue("SIGN", "I");
                rtSourceDocDate.SetValue("OPTION", "BT");
                rtSourceDocDate.SetValue("LOW", firstDate.ToString("yyyyMMdd"));
                rtSourceDocDate.SetValue("HIGH", lastDate.ToString("yyyyMMdd"));

                rfcFun.Invoke(rfcDest);
                IRfcTable rtSourceOuput = rfcFun.GetTable("Z_OUTPUT");

                DataTable dtDN = ConvertToDoNetTable(rtSourceOuput);
                return dtDN;
            }
            catch (RfcCommunicationException ex)
            {
                WriteLogs.Write("DownloadDeliverySale", ex);
                return null;
            }
            catch (RfcAbapException ex)
            {
                WriteLogs.Write("DownloadDeliverySale", ex);
                return null;
            }
            catch (RfcLogonException ex)
            {
                WriteLogs.Write("DownloadDeliverySale", ex);
                return null;
            }
            catch (Exception ex)
            {
                WriteLogs.Write("DownloadDeliverySale", ex);
                return null;
            }
        }

        public DataTable DownloadTODetail(string toNumber)
        {
            try
            {
                RfcConfigParameters rfcParam = Connector3SapRfcConnGroup(_SAPConnectionString);
                RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcParam);
                RfcRepository rfcRep = rfcDest.Repository;
                IRfcFunction rfcFun = rfcRep.CreateFunction("BAPI_WHSE_TO_GET_DETAIL");//SAP RFC
                rfcFun.SetValue("WHSENUMBER", "049");
                rfcFun.SetValue("TRANSFERORDERNO", toNumber);
                rfcFun.SetValue("TRANSFERORDERITEM", "");

                rfcFun.Invoke(rfcDest);
                IRfcTable rtSource = rfcFun.GetTable("TOITEMDATA");
                DataTable dtMaterialMaster = ConvertToDoNetTable(rtSource);
                return dtMaterialMaster;
            }
            catch (RfcCommunicationException ex)
            {
                WriteLogs.Write("DownloadTODetail", ex);
                return null;
            }
            catch (RfcAbapException ex)
            {
                WriteLogs.Write("DownloadTODetail", ex);
                return null;
            }
            catch (RfcLogonException ex)
            {
                WriteLogs.Write("DownloadTODetail", ex);
                return null;
            }
            catch (Exception ex)
            {
                WriteLogs.Write("DownloadTODetail", ex);
                return null;
            }
        }

        public PartInfo GetAvailableStock(string line, string part)
        {
            try
            {
                RfcConfigParameters rfcParam = Connector3SapRfcConnGroup(_SAPConnectionString);
                RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcParam);
                RfcRepository rfcRep = rfcDest.Repository;

                IRfcFunction rfcFun = rfcRep.CreateFunction("ZRFC_WM_STOCK_0001");//SAP RFC

                rfcFun.SetValue("I_LGNUM", "049");
                rfcFun.SetValue("I_LGTYP", line);
                rfcFun.SetValue("I_WERKS", "");
                rfcFun.SetValue("I_LGORT", "");
                rfcFun.SetValue("I_MATNR", part);

                rfcFun.Invoke(rfcDest);

                IRfcTable rtSource = rfcFun.GetTable("TAB_LQUA");
                DataTable dtMaterialMaster = ConvertToDoNetTable(rtSource);
                if (dtMaterialMaster == null || dtMaterialMaster.Rows.Count > 0)
                {
                    int qty = int.TryParse(dtMaterialMaster.Rows[0]["VERME"].ToString().Split('.')[0], out int _qty) ? _qty : -1;
                    string plant = dtMaterialMaster.Rows[0]["WERKS"].ToString();
                    return new PartInfo
                    {
                        Plant = plant,
                        Qty = qty
                    };
                }
                else
                {
                    return null;
                }
            }
            catch (RfcCommunicationException ex)
            {
                WriteLogs.Write("GetAvailableStock", ex);
                return null;
            }
            catch (RfcAbapException ex)
            {
                WriteLogs.Write("GetAvailableStock", ex);
                return null;
            }
            catch (RfcLogonException ex)
            {
                WriteLogs.Write("GetAvailableStock", ex);
                return null;
            }
            catch (Exception ex)
            {
                WriteLogs.Write("GetAvailableStock", ex);
                return null;
            }
        }

        public DataTable GetListTO(DateTime from, DateTime to)
        {
            try
            {
                RfcConfigParameters rfcParam = Connector3SapRfcConnGroup(_SAPConnectionString);
                RfcDestination rfcDest = RfcDestinationManager.GetDestination(rfcParam);
                RfcRepository rfcRep = rfcDest.Repository;

                IRfcFunction rfcFun = rfcRep.CreateFunction("ZRFC_MM_GET_TOLIST");//SAP RFC

                string _hFrom = from.ToString("HHmmss");
                string _hTo = to.ToString("HHmmss");

                rfcFun.SetValue("I_LGNUM", "049");
                rfcFun.SetValue("I_BDATUF", from); //from date
                rfcFun.SetValue("I_BDATUT", to); //to date
                rfcFun.SetValue("I_BZEITF", _hFrom); //from time
                rfcFun.SetValue("I_BZEITT", _hTo); // to time

                rfcFun.Invoke(rfcDest);

                IRfcTable rtSource = rfcFun.GetTable("TOLIST");
                DataTable dtMaterialMaster = ConvertToDoNetTable(rtSource);
                if (dtMaterialMaster == null || dtMaterialMaster.Rows.Count > 0)
                {
                    var data = new DataTable();
                    data.Columns.Add("ID", System.Type.GetType("System.Int64")).SetOrdinal(0);

                    var column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "TO";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("	System.DateTime");
                    column.ColumnName = "CREATED_DATE";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.TimeSpan");
                    column.ColumnName = "CREATED_TIME";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "TO_ITEM";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "COMPONENT";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "SOURCE";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "LINE";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.Int32");
                    column.ColumnName = "QTY";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "PON";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "BIN";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "MATERIAL_DESCRIPTION";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "UNIT";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "PICK_AREA";
                    data.Columns.Add(column);

                    column = new DataColumn();
                    column.DataType = System.Type.GetType("System.String");
                    column.ColumnName = "PLANT";
                    data.Columns.Add(column);


                    foreach (DataRow item in dtMaterialMaster.Rows)
                    {
                        if (!string.IsNullOrEmpty(item["TANUM"].ToString()) /*&& balines.Contains(item["NLTYP"].ToString())*/ /*&& item["VLTYP"].ToString().Equals("BFO")*/)
                        {
                            DataRow row = data.NewRow();
                            row[1] = item["TANUM"].ToString();
                            row[2] = item["BDATU"].ToString();
                            row[3] = item["BZEIT"].ToString();
                            row[4] = item["TAPOS"].ToString();
                            row[5] = item["MATNR"].ToString();
                            row[6] = item["VLTYP"].ToString();
                            row[7] = item["NLTYP"].ToString();
                            row[8] = int.TryParse(item["NSOLM"].ToString().Split('.')[0], out int _qty) ? _qty : -1;
                            row[9] = item["BENUM"].ToString();
                            row[10] = item["VLPLA"].ToString();
                            row[11] = item["MAKTX"].ToString();
                            row[12] = item["MEINS"].ToString();
                            row[13] = item["KOBER"].ToString();
                            row[14] = item["WERKS"].ToString();

                            data.Rows.Add(row);
                        }
                        return data;
                    }
                }
                return null;
            }
            catch (RfcCommunicationException ex)
            {
                WriteLogs.Write("GetListTO", ex);
                return null;
            }
            catch (RfcAbapException ex)
            {
                WriteLogs.Write("GetListTO", ex);
                return null;
            }
            catch (RfcLogonException ex)
            {
                WriteLogs.Write("GetListTO", ex);
                return null;
            }
            catch (Exception ex)
            {
                WriteLogs.Write("GetListTO", ex);
                return null;
            }
        }


    }

    public class PartInfo
    {
        public string Plant { get; set; }
        public int Qty { get; set; }
    }
    public class DataVM
    {
        public long TO { get; set; }
        public int TO_ORDER { get; set; }
        public string BIN { get; set; }
        public string MATERIAL { get; set; }
        public int QTY { get; set; }
        public string UNIT { get; set; }
    }
}