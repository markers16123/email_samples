// See https://aka.ms/new-console-template for more information

Console.WriteLine("Hello, World!");

string sFile = @"C:\temp\39@1516436233049.eml";
try
{
    CDO.IMessage message = ReadMessage(sFile);
    CDO.IBodyParts attachments = message.Attachments;
    int count = attachments.Count;
    foreach(CDO.IBodyPart attachment in attachments)
    {
        ADODB.Stream stream = attachment.GetDecodedContentStream();
        Console.WriteLine($"{attachment.FileName} - {stream.Size}");
        stream.SaveToFile($@"C:\temp\39@1516436233049.eml.{attachment.FileName}", ADODB.SaveOptionsEnum.adSaveCreateOverWrite);
    }
}
catch (System.IO.IOException err)
{
    Console.WriteLine("File " + sFile + " is currently in use.");
    Console.WriteLine(err);
}

CDO.Message ReadMessage(String emlFileName)
{
    CDO.Message msg = new CDO.Message();
    ADODB.Stream stream = new ADODB.Stream();
    stream.Open(Type.Missing,
                   ADODB.ConnectModeEnum.adModeUnknown,
                   ADODB.StreamOpenOptionsEnum.adOpenStreamUnspecified,
                   String.Empty,
                   String.Empty);
    stream.LoadFromFile(emlFileName);
    stream.Flush();
    msg.DataSource.OpenObject(stream, "_Stream");
    msg.DataSource.Save();
    return msg;
}