using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GPRSTOOL
{

    public class GPRSparam : ICloneable 
    {
        string country = "";//国家

        public string Country
        {
            get { return country; }
            set { country = value; }
        }
        string typestr = "";

        public string Typestr
        {
            get { return typestr; }
            set { typestr = value; }
        }
        GPRSTYPE type = GPRSTYPE.GPRS;


        public GPRSTYPE Type
        {
            get { return type; }
            set { type = value; }
        }
        string mvno_type = "";//虚拟运营商类型  spn
        
        public string Mvno_type
        {
            get { return mvno_type; }
            set { mvno_type = value; }
        }
        string mvno_match_data = "";//虚拟运营商类型值 MTN Tigo Airtel

        public string Mvno_match_data
        {
            get { return mvno_match_data; }
            set { mvno_match_data = value; }
        }
        string idledisplay = "";

        public string Idledisplay
        {
            get { return idledisplay; }
            set { idledisplay = value; }
        }
        string operatorName = "";

        public string OperatorName
        {
            get { return operatorName; }
            set { operatorName = value; }
        }

        string homepage = "";
        public string Homepage
        {
          get { return homepage; }
          set { homepage = value; }
        }

        string name = "";

        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        string apn = "";

        public string Apn
        {
            get { return apn; }
            set { apn = value; }
        }
        string proxyEnable ="";
        public string ProxyEnable
        {
          get { return proxyEnable; }
          set { proxyEnable = value; }
        }
        string proxy = "";

        public string Proxy
        {
            get { return proxy; }
            set { proxy = value; }
        }
        string port = "";

        public string Port
        {
            get { return port; }
            set { port = value; }
        }
        string username = "";

        public string Username
        {
            get { return username; }
            set { username = value; }
        }
        string password = "";

        public string Password
        {
            get { return password; }
            set { password = value; }
        }
        string server = "";

        public string Server
        {
            get { return server; }
            set { server = value; }
        }
        string mmsc = "";

        public string Mmsc
        {
            get { return mmsc; }
            set { mmsc = value; }
        }
        string mmsproxy = "";

        public string Mmsproxy
        {
            get { return mmsproxy; }
            set { mmsproxy = value; }
        }
        string mmsport = "";

        public string Mmsport
        {
            get { return mmsport; }
            set { mmsport = value; }
        }
        string mcc = "";

        public string Mcc
        {
            get { return mcc; }
            set { mcc = value; }
        }
        string mnc = "";

        public string Mnc
        {
            get { return mnc; }
            set { mnc = value; }
        }
        string authtype = "";

        public string Authtype
        {
            get { return authtype; }
            set { authtype = value; }
        }
        string apntype = "";

        public string Apntype
        {
            get { return apntype; }
            set { apntype = value; }
        }

        public object Clone()
        {
            return this.MemberwiseClone(); 
        }

    }
    public enum GPRSTYPE
    {
        GPRS,
        MMS
    }
}
