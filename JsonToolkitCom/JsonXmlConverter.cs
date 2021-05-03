using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace JsonToolkitCom
{

    [ComVisible(true)]
    [Guid("7FD0B31A-8DAD-4F87-B325-16789EBB985E")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IJsonXmlConverter
    {
        string XmlToJson(string json);
        string JsonToXml(string xml);
    }

    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    [Guid("1012B3FF-333B-43CA-A931-A9EAB4A4643E")]
    [ProgId("JsonToolkitCom.JsonXmlConverter")]
    [ComDefaultInterface(typeof(IJsonXmlConverter))]
    public class JsonXmlConverter : IJsonXmlConverter
    {
        public string JsonToXml(string xml)
        {
            return JsonConvert.SerializeXNode(XDocument.Parse(xml).Root);
        }

        public string XmlToJson(string json)
        {
            return JsonConvert.DeserializeXNode(json).ToString();
        }
    }
}
