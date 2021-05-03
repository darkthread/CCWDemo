using Newtonsoft.Json;
using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
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
        public string XmlToJson(string xml)
        {
            return JsonConvert.SerializeXNode(XDocument.Parse(xml).Root);
        }

        public string JsonToXml(string json)
        {
            return JsonConvert.DeserializeXNode(json).ToString();
        }

        static JsonXmlConverter()
        {
            AppDomain.CurrentDomain.AssemblyResolve += OnResolveAssembly;
        }

        private static Assembly OnResolveAssembly(object sender, ResolveEventArgs args)
        {
            Assembly executingAssembly = Assembly.GetExecutingAssembly();
            AssemblyName assemblyName = new AssemblyName(args.Name);

            string path = assemblyName.Name + ".dll";
            if (assemblyName.CultureInfo.Equals(CultureInfo.InvariantCulture) == false)
            {
                path = String.Format(@"{0}\{1}", assemblyName.CultureInfo, path);
            }

            using (Stream stream = executingAssembly.GetManifestResourceStream(path))
            {
                if (stream == null)
                    return null;

                byte[] assemblyRawBytes = new byte[stream.Length];
                stream.Read(assemblyRawBytes, 0, assemblyRawBytes.Length);
                return Assembly.Load(assemblyRawBytes);
            }
        }

    }


}
