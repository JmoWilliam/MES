using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.IO;
using System.Web;

public class UnicodeFontFactory : FontFactoryImp
{
    private static readonly string arialFontPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "ARIALUNI.TTF"); //arial unicode MS是完整的unicode字型。
    private static readonly string 標楷體Path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "KAIU.TTF"); //標楷體
    private static readonly string 新細明體Path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "mingliu.ttc,1"); //新細明體

    public override Font GetFont(string fontname, string encoding, bool embedded, float size, int style, BaseColor color, bool cached)
    {
        //可用Arial或標楷體，自己選一個
        BaseFont baseFont = BaseFont.CreateFont(新細明體Path, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        return new Font(baseFont, size, style, color);
    }
}