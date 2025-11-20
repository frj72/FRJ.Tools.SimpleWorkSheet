using FRJ.Tools.SimpleWorkSheet.Components.Book;
using FRJ.Tools.SimpleWorkSheet.Components.Sheet;

namespace FRJ.Tools.SimpleWorkSheet.Showcase.Examples.EdgeCases;

public class UnicodeExample : IShowcase
{
    public string Name => "Unicode Characters (Emoji, Arabic, Chinese, Special chars)";
    public string Description => "Tests handling of various Unicode characters including RTL text";
    public string Category => "Edge Cases";

    public void Run()
    {
        var sheet = new WorkSheet("Unicode");
        
        sheet.AddCell(new(0, 0), "Type", cell => cell.WithFont(f => f.Bold()));
        sheet.AddCell(new(1, 0), "Text", cell => cell.WithFont(f => f.Bold()));
        
        sheet.AddCell(new(0, 1), "Emoji", null);
        sheet.AddCell(new(1, 1), "😀 😃 😄 😁 🎉 🎊 ✨ 🌟 ⭐ 🚀 🔥 💯", null);
        
        sheet.AddCell(new(0, 2), "Arabic (RTL)", null);
        sheet.AddCell(new(1, 2), "مرحبا بك في عالم البرمجة", null);
        
        sheet.AddCell(new(0, 3), "Chinese", null);
        sheet.AddCell(new(1, 3), "欢迎使用简单工作表库", null);
        
        sheet.AddCell(new(0, 4), "Japanese", null);
        sheet.AddCell(new(1, 4), "こんにちは世界", null);
        
        sheet.AddCell(new(0, 5), "Korean", null);
        sheet.AddCell(new(1, 5), "안녕하세요 세계", null);
        
        sheet.AddCell(new(0, 6), "Mathematical", null);
        sheet.AddCell(new(1, 6), "∑ ∫ ∂ √ ∞ ≈ ≠ ≤ ≥ ± ×  ÷", null);
        
        sheet.AddCell(new(0, 7), "Currency", null);
        sheet.AddCell(new(1, 7), "$ € £ ¥ ₹ ₽ ₿ ¢", null);
        
        sheet.AddCell(new(0, 8), "Diacritics", null);
        sheet.AddCell(new(1, 8), "àáâãäå èéêë ìíîï òóôõö ùúûü ñ ç", null);
        
        sheet.SetColumnWidth(1, 50.0);
        
        var workbook = new WorkBook("Unicode", [sheet]);
        ShowcaseRunner.SaveWorkBook(workbook, "Showcase_03_Unicode.xlsx");
    }
}
