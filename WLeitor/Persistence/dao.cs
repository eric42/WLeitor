using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.OracleClient;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Npgsql;

namespace WLeitor.ECD
{
    class dao
    {
        public static string strConn;
        public List<string> dados = new List<string>();

        /*
        private string ProviderName;
        private DbProviderFactory provider;
        private string pEngine;
        private DbConnection connSQL;
        DbTransaction oTrans;
        private string prefixParameter;
        private string prefixParamSelect;
        private string pFuncaoData;
        private string pFuncaoNull;
        private bool disposed = false;

        public Banco()
        {
            strConn = Configuracao.strConexao;
            ProviderName = Configuracao.ProviderName;
            pEngine = Configuracao.ProviderName == "System.Data.OracleClient" ? "Oracle" : "MSSQL";
            pEngine = pEngine.ToUpper();
            switch (pEngine)
            {
                case "MSSQL":
                    prefixParameter = "@";
                    prefixParamSelect = "@";
                    pFuncaoData = "getdate()";
                    pFuncaoNull = "isnull";
                    break;
                case "ORACLE":
                    prefixParameter = "p_";
                    prefixParamSelect = ":p_";
                    pFuncaoData = "sysdate";
                    pFuncaoNull = "nvl";
                    break;
                default:
                    ProviderName = "System.Data.SqlClient";
                    break;
            }

            provider = DbProviderFactories.GetFactory(ProviderName);

            //connSQL = new SqlConnection(strConexao);
            connSQL = provider.CreateConnection();
            connSQL.ConnectionString = strConn;
        }*/

        public static string getStrConn(int pIndice, string pLogin, string pSenha)
        {
            switch (pIndice)
            {
                case 1:
                    //strConn = "Provider=MSDAORA;Data Source=SSVMBDI08\\SSVMBDI08;Initial Catalog=ABBOTT;User ID=" + pLogin+";Password="+pSenha+";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI08\\SSVMBDI08;Port=1523;Database=ABBOTT;User ID=" + pLogin+";Password="+pSenha+ ";Trusted_Connection=True;";
                    //strConn = "Provider=MSDAORA;Data Source=SSVMBDI08\\SSVMBDI08;location=SSVMBDI08\\SSVMBDI08;User ID=" + pLogin + ";password=" + pSenha + ";timeout=1000;";
                    strConn = "Provider = OraOLEDB.Oracle; Data Source = ABBOTT; User ID= = " + pLogin + "; Password = " + pSenha + ";";
                    break;
                case 2:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI08\\SSVMBDI08;Initial Catalog=ABBVIE;User ID="+pLogin+ "; Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI08\\SSVMBDI08;Port=1526;User ID=" + pLogin + "; Password=" + pSenha + ";Database=ABBVIE;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI08\\SSVMBDI08; DATABASE = ABBVIE; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 3:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI08\\SSVMBDI08;Initial Catalog=BPOAPP;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI08\\SSVMBDI08;Port=1535;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BPOAPP;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI08\\SSVMBDI08; DATABASE = BPOAPP; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 4:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI08\\SSVMBDI08;Initial Catalog=NOVARTIS;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI08\\SSVMBDI08;Port=1524;User ID=" + pLogin + ";Password=" + pSenha + ";Database=NOVARTIS;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI08\\SSVMBDI08; DATABASE = NOVARTIS; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 5:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI08\\SSVMBDI08;Initial Catalog=REFARMA;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI08\\SSVMBDI08;Port=1522;User ID=" + pLogin + ";Password=" + pSenha + ";Database=REFARMA;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI08\\SSVMBDI08; DATABASE = REFARMA; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 6:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI08\\SSVMBDI08;Initial Catalog=SUZANO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI08\\SSVMBDI08;Port=1521;User ID=" + pLogin + ";Password=" + pSenha + ";Database=SUZANO;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI08\\SSVMBDI08; DATABASE = SUZANO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 7:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=ABBOTTCL;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1540;User ID=" + pLogin + ";Password=" + pSenha + ";Database=abbottcl;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = abbottcl; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 8:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=AJIHANA;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1529;User ID=" + pLogin + ";Password=" + pSenha + ";Database=AJIHANA;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = AJIHANA; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 9:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=ACHEMA;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1522;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ACHEMA;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = ACHEMA; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 10:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=BAUDUCCO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1531;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BAUDUCCO;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = BAUDUCCO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 11:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=BEMIS;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1524;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BEMIS;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = BEMIS; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 12:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=BENTELER;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1525;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BENTELER;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = BENTELER; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 13:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=BRF;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1532;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BRF;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = BRF; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 14:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=CENTAURO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1537;User ID=" + pLogin + ";Password=" + pSenha + ";Database=CENTAURO;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = CENTAURO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 15:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=CONECTA;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1526;User ID=" + pLogin + ";Password=" + pSenha + ";Database=CONECTA;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = CONECTA; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 16:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=CONTPRD;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1521;User ID=" + pLogin + ";Password=" + pSenha + ";Database=CONTPRD;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = CONTPRD; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 17:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=DBACORP;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1515;User ID=" + pLogin + ";Password=" + pSenha + ";Database=DBACORP;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = DBACORP; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 18:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=EDOCFLOW;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1528;User ID=" + pLogin + ";Password=" + pSenha + ";Database=EDOCFLOW;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = EDOCFLOW; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 19:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=EDOCRECO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1523;User ID=" + pLogin + ";Password=" + pSenha + ";Database=EDOCRECO;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = EDOCRECO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 20:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=IMMB;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1527;User ID=" + pLogin + ";Password=" + pSenha + ";Database=IMMB;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = IMMB; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 21:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=MEDTONIC;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1534;User ID=" + pLogin + ";Password=" + pSenha + ";Database=MDTRONIC;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = MEDTONIC; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 22:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=PHILIP;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1535;User ID=" + pLogin + ";Password=" + pSenha + ";Database=PHILIP;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = PHILIP; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 23:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=SARALEE;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1530;User ID=" + pLogin + ";Password=" + pSenha + ";Database=SARALEE;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = SARALEE; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 24:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI09\\SSVMBDI09;Initial Catalog=SATTOIL;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1536;User ID=" + pLogin + ";Password=" + pSenha + ";Database=SATTOIL;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = SATTOIL; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 25:
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI09\\SSVMBDI09;Port=1533;User ID=" + pLogin + ";Password=" + pSenha + ";Database=WHIRL;";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=localhost;Port=1533;User ID=" + pLogin + ";Password=" + pSenha + ";Database=WHIRL;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI09\\SSVMBDI09; DATABASE = WHIRL; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 26:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI10\\SSVMBDI10;Initial Catalog=ALLERGAN;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI10\\SSVMBDI10;Port=1524;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ALLERGAN;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI10\\SSVMBDI10; DATABASE = ALLERGAN; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 27:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI10\\SSVMBDI10;Initial Catalog=CPKELCO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI10\\SSVMBDI10;Port=1526;User ID=" + pLogin + ";Password=" + pSenha + ";Database=CPKELCO;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI10\\SSVMBDI10; DATABASE = CPKELCO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 28:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI10\\SSVMBDI10;Initial Catalog=ELECTRO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI10\\SSVMBDI10;Port=1536;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ELECTRO;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI10\\SSVMBDI10; DATABASE = ELECTRO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 29:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI10\\SSVMBDI10;Initial Catalog=ENAMEL;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI10\\SSVMBDI10;Port=1527;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ENAMEL;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI10\\SSVMBDI10; DATABASE = ENAMEL; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 30:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI10\\SSVMBDI10;Initial Catalog=FMCT;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI10\\SSVMBDI10;Port=1529;User ID=" + pLogin + ";Password=" + pSenha + ";Database=FMCT;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI10\\SSVMBDI10; DATABASE = FMCT; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 31:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=BOSTON;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1521;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BOSTON;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = BOSTON; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                /*case 32:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=BOSTON;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1521;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BOSTON;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = BOSTON; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;*/
                case 32:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=BRFTEMP;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1534;User ID=" + pLogin + ";Password=" + pSenha + ";Database=BRFTEMP;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = BRFTEMP; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 34:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=CONTCMPL;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1523;User ID=" + pLogin + ";Password=" + pSenha + ";Database=CONTCMPL;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = CONTCMPL; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 35:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=EVONIK;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1527;User ID=" + pLogin + ";Password=" + pSenha + ";Database=EVONIK;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = EVONIK; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 36:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=FLOWSERVER;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1528;User ID=" + pLogin + ";Password=" + pSenha + ";Database=FSERVE;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = FLOWSERVER; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 37:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=IPIRANGA;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1530;User ID=" + pLogin + ";Password=" + pSenha + ";Database=IPIRANGA;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = IPIRANGA; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 38:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=LOCAWEB;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1535;User ID=" + pLogin + ";Password=" + pSenha + ";Database=LOCAWEB;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = LOCAWEB; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 39:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=NOVOMG;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1525;User ID=" + pLogin + ";Password=" + pSenha + ";Database=NOVOMG;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = NOVOMG; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 40:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=NOVONORD;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1526;User ID=" + pLogin + ";Password=" + pSenha + ";Database=NOVONORD;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = NOVONORD; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 41:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=RFARMA;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1529;User ID=" + pLogin + ";Password=" + pSenha + ";Database=RFARMA;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = RFARMA; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 42:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=REPSOL;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1537;User ID=" + pLogin + ";Password=" + pSenha + ";Database=REPSOL;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = REPSOL; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 43:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=SIN;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1524;User ID=" + pLogin + ";Password=" + pSenha + ";Database=SIN;";
                    break;
                case 44:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=TORRENT;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1531;User ID=" + pLogin + ";Password=" + pSenha + ";Database=TORRENT;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = TORRENT; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 45:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=VILLARES;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1532;User ID=" + pLogin + ";Password=" + pSenha + ";Database=VILLARES;";
                    break;
                case 46:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSVMBDI11\\SSVMBDI11;Initial Catalog=WARTSILA;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSVMBDI11\\SSVMBDI11;Port=1533;User ID=" + pLogin + ";Password=" + pSenha + ";Database=WARTSILA;";
                    strConn = "Provider = MSDAORA; Data Source= SSVMBDI11\\SSVMBDI11; DATABASE = WARTSILA; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 47:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=172.18.44.36\\172.16.48.36;Initial Catalog=SATP;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=172.18.44.36\\172.16.48.36;Port=1529;User ID=" + pLogin + ";Password=" + pSenha + ";Database=SATP;";
                    strConn = "Provider = MSDAORA; Data Source= 172.18.44.36\\172.16.48.36; DATABASE = SATP; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 48:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=10.140.48.76\\10.140.48.76;Initial Catalog=LFT;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=10.140.48.76\\10.140.48.76;Port=1521;User ID=" + pLogin + ";Password=" + pSenha + ";Database=LFT;";
                    strConn = "Provider = MSDAORA; Data Source= 10.140.48.76\\10.140.48.76; DATABASE = LFT; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 49:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Initial Catalog=ANGLO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Port=1528;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ANGLO;";
                    strConn = "Provider = MSDAORA; Data Source= 10.140.48.76\\10.140.48.76; DATABASE = ANGLO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 50:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Initial Catalog=ANGLO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Port=1531;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ANGLO;";
                    strConn = "Provider = MSDAORA; Data Source= SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br; DATABASE = ANGLO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 51:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Initial Catalog=ANGLO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Port=1530;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ANGLO;";
                    strConn = "Provider = MSDAORA; Data Source= SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br; DATABASE = ANGLO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 52:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Initial Catalog=ANGLO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Port=1550;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ANGLO;";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Port=1550;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ANGLO;";
                    strConn = "Provider = MSDAORA; Data Source= SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br; DATABASE = ANGLO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                case 53:
                    //strConn = "Provider=OraOLEDB.Oracle;Data Source=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Initial Catalog=ANGLO;User ID=" + pLogin + ";Password=" + pSenha + ";";
                    //strConn = "Provider=OraOLEDB.Oracle;Server=SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br;Port=1529;User ID=" + pLogin + ";Password=" + pSenha + ";Database=ANGLO;";
                    strConn = "Provider = MSDAORA; Data Source= SSBDISP03.sondait.com.br\\SSBDISP03.sondait.com.br; DATABASE = ANGLO; UID = " + pLogin + "; PWD = " + pSenha + ";";
                    break;
                default:
                    strConn = "";
                    break;
            }

            return strConn;
        }

        public OleDbConnection GetConection()
        {
            //return new NpgsqlConnection(strConn);

            return new OleDbConnection(strConn);
        }

        public NpgsqlConnectionFactory Conection()
        {
            NpgsqlConnectionFactory conn = new NpgsqlConnectionFactory();
            conn.CreateConnection(strConn);
            return conn;
        }

        public DataTable executaComando()
        {
            /* var test = new NpgsqlConnectionFactory();
             dbCommand cmd = new dbCommand();*/
            OleDbConnection cn = new OleDbConnection();
            OleDbCommand dbCommand = cn.CreateCommand();
            DataTable oDt = new DataTable();

            cn = GetConection();

            string sql = "select ent.id_nf_entrada, ent.chave_nf, ent.num_nf, ent.dt_emissao," +
                " ent.dt_entrada, ent.cgc_cpf, itm.id_item, itm.cod_produto, itm.descricao_nota," +
                " itm.cod_unid_medida, itm.quantidade, itm.vlr_bruto" +
                "from lf_nf_entrada ent, lf_nf_entrada_item" +
                "where ent.cod_matriz = '" + "'" +
                "and ent.cod_filial = '" + "'" +
                "and ent.dt_entrada between '" + "' and '" + "'" +
                "and.ent.cod_modelo = '55'" +
                "and ent.cod_status = '01'" +
                "and ent.cod_matriz = itm.cod_matriz" +
                "and ent.cod_filial = itm.cod_filial" +
                "and ent.id_nf_entrada = itm.id_nf_entrada;";

            dbCommand.CommandText = sql;

            try
            {
                cn.Open();
                dbCommand.Connection = cn;
                OracleDataAdapter oDa = new OracleDataAdapter();
                oDa.Fill(oDt);

                return oDt;
            }
            catch (OleDbException ex)
            {
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
                dbCommand.Dispose();
                cn.Dispose();

                StreamWriter pod;
                string caminho, path;

                caminho = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                path = caminho + @"\log.txt";
                pod = File.CreateText(path);

                pod.WriteLine(ex.Message);
                pod.Flush();
                throw (ex);
                //return null;
            }
            finally
            {
                if (cn.State == ConnectionState.Open)
                {
                    cn.Close();
                }
                dbCommand.Dispose();
                cn.Dispose();

            }
        }

        public void ConectaBanco(string pUser, string pPass, string pBanco, string pMatriz, string pFilial, string pDtIni, string pDtFin)
        {
            string erro;
            string user = pUser;
            string pass = pPass;
            //String strConn = "Provider=MSDAORA;Data Source=" + pBanco + ";User ID=" + user + ";Password=" + pass 
            String strConn = "Provider=OraOLEDB.Oracle;Data Source=" + pBanco + ";User ID=" + user + ";Password=" + pass;
            OleDbConnection cnn = new OleDbConnection(strConn);
            OleDbCommand cmd;
            DataTable dt = new DataTable();

            try
            {
                cnn.Open();

                string sql = "select ent.id_nf_entrada, ent.chave_nf, ent.num_nf, ent.dt_emissao," +
                " ent.dt_entrada, ent.cgc_cpf, itm.id_item, itm.cod_produto, itm.descricao_nota," +
                " itm.cod_unid_medida, itm.quantidade, itm.vlr_bruto" +
                "from lf_nf_entrada ent, lf_nf_entrada_item" +
                "where ent.cod_matriz = '" + pMatriz + "'" +
                "and ent.cod_filial = '" + pFilial + "'" +
                "and ent.dt_entrada between '" + pDtIni + "' and '" + pDtFin + "'" +
                "and.ent.cod_modelo = '55'" +
                "and ent.cod_status = '01'" +
                "and ent.cod_matriz = itm.cod_matriz" +
                "and ent.cod_filial = itm.cod_filial" +
                "and ent.id_nf_entrada = itm.id_nf_entrada;";

                cmd = new OleDbCommand(sql, cnn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                da = null;
                cnn.Close();

            }
            catch (Exception ex)
            {
                if (ex.Message == "ORA-12541: TNS:não há listener")
                {
                    MessageBox.Show("Banco inválido!", "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show(ex.Message, "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            finally
            {
                cnn = null;
                cmd = null;
                dt = null;
            }
        }

        public ArrayList carregaCombo(string TNSNAMES)
        {
            string sTNSNames = TNSNAMES;
            string sDir = "";
            string s_TnsNames = "";
            ArrayList DBNamesCollection = new ArrayList();
            string RegExPattern = @"[\n][\s]*[^\(][a-zA-Z0-9_.]+[\s]*=[\s]*\(";
            int iCount = 0;
            frmParametros frm = new frmParametros();

            if (System.IO.File.Exists(sTNSNames) != true)
            {
                MessageBox.Show("TNSNAMES não está nesse dir: " + sDir + "", "Erro", MessageBoxButtons.OK);
                throw new System.IO.FileNotFoundException("TNSNAMES não está nesse dir: " + sDir + "");
            }

            DBNamesCollection.Add("");

            if (!string.IsNullOrEmpty(TNSNAMES))
            {
                try
                {
                    var fiTNS = new System.IO.FileInfo(sTNSNames);
                    if (fiTNS.Exists)
                    {
                        if (fiTNS.Length > 0)
                        {

                            try
                            {
                                int countFin = System.Text.RegularExpressions.Regex.Matches(System.IO.File.ReadAllText(fiTNS.FullName), RegExPattern).Count - 1;
                                for (iCount = 0; iCount <= countFin; iCount++)
                                {
                                    s_TnsNames = Regex.Matches(File.ReadAllText(fiTNS.FullName), RegExPattern)[iCount].Value.Trim().Substring(0, Regex.Matches(File.ReadAllText(fiTNS.FullName), RegExPattern)[iCount].Value.Trim().IndexOf(" ")).Trim();
                                    if (s_TnsNames != "")
                                        DBNamesCollection.Add(s_TnsNames);
                                }

                            }
                            catch (Exception ex)
                            {
                                throw new Exception(ex.Message);
                            }
                            
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
            }
            DBNamesCollection.Sort();
            frm.sRetornoErro = "certo";
            return DBNamesCollection;
        }

        private string FRetiraCaracter(string sTnsNames)
        {
            int iTam = sTnsNames.Length;
            string sTnsNameLimpo = "";
            string sCaracter = "";
            int i = 0;
            for (i = 0; i == (iTam - 1); i++)
            {
                sCaracter = sTnsNames.Substring(i, 1).Trim().ToString();
                if (sCaracter.Contains("[A-Za-z0-9-_]"))
                    sTnsNameLimpo = sTnsNameLimpo + sCaracter;
            }
            return sTnsNameLimpo;
        }

        public static void carregaBanco(ComboBox cbo, string strsql, string sBanco, string sRet, string sUser, string sPwd)
        {
            sRet = "";
            string strConn = "Provider=OraOLEDB.Oracle;Data Source=" + sBanco + ";User ID=" + sUser + ";Password=" + sPwd;
            OleDbConnection cnn = new OleDbConnection(strConn);
            var cmd = new OleDbCommand();
            var dt = new DataTable();
            frmParametros frm = new frmParametros();
            frm.oDt = null;

            try
            {
                cnn.Open();
                cmd = new OleDbCommand(strsql, cnn);
                var da = new OleDbDataAdapter(cmd);
                da.Fill(dt);
                da.Fill(frm.oDt);

                cbo.DataSource = null;
                cbo.DataSource = dt;
                cbo.ValueMember = dt.Columns[0].ToString();
                cbo.DisplayMember = dt.Columns[0].ToString();

                da = null;
                cnn.Close();

            }
            catch (Exception ex)
            {
                if (ex.Message == "ORA-12541: TNS:não há listener")
                {
                    MessageBox.Show("Banco inválido!", "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cbo.DataSource = null;
                    cbo.Items.Clear();
                    frm.sRetornoErro = ex.Message;
                }
                else
                {
                    MessageBox.Show(ex.Message, "Mensagem", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    sRet = ex.Message;
                    frm.sRetornoErro = ex.Message;
                }
            }
            finally
            {
                dt = null;
                cnn = null;
                cmd = null;

            }

        }
    }
}

