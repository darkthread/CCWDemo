Set conv = CreateObject("JsonToolkitCom.JsonXmlConverter")
WScript.Echo conv.XmlToJson("<User><Id>A001</Id><Name>Jeffrey</Name></User>")
WScript.Echo conv.JsonToXml("{ ""User"": { ""Id"": ""D001"", ""Name"": ""darkthread"" } }")
