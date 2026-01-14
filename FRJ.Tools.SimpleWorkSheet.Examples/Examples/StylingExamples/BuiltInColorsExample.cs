using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;
using FRJ.Tools.SimpleWorkSheet.Examples.Examples.Utils;

namespace FRJ.Tools.SimpleWorkSheet.Examples.Examples.StylingExamples;

public class BuiltInColorsExample : IExample
{
    public string Name => "078_BuiltInColors";
    public string Description => "Showcase of all built-in colors with visual samples";

    public void Run()
    {
        var sheet = new WorkSheet("Built-in Colors");
        
        var colorList = GetAllColors();

        AddHeaders(sheet);
        AddColorRows(sheet, colorList);

        sheet.SetColumnWidth(0, 5.0);
        sheet.SetColumnWidth(1, 30.0);
        sheet.FreezePanes(1, 0);

        ExampleRunner.SaveWorkSheet(sheet, $"{Name}.xlsx");
    }

    private static void AddHeaders(WorkSheet sheet)
    {
        sheet.AddCell(new(0, 0), new("Color"), builder => builder
            .WithStyle(style => style
                .WithFont(font => font.Bold().WithSize(12))
                .WithFillColor(Colors.DarkBlue)
                .WithFont(font => font.WithColor(Colors.White))));

        sheet.AddCell(new(1, 0), new("Name"), builder => builder
            .WithStyle(style => style
                .WithFont(font => font.Bold().WithSize(12))
                .WithFillColor(Colors.DarkBlue)
                .WithFont(font => font.WithColor(Colors.White))));
    }

    private static void AddColorRows(WorkSheet sheet, Dictionary<string, string> colors)
    {
        uint row = 1;
        foreach (var color in colors)
        {
            sheet.AddCell(new(0, row), new(" "), builder => builder
                .WithStyle(style => style.WithFillColor(color.Value)));
            
            sheet.AddCell(new(1, row), new(color.Key), null);
            
            row++;
        }
    }

    private static Dictionary<string, string> GetAllColors() => new()
    {
        { "Black", Colors.Black },
        { "White", Colors.White },
        { "Red", Colors.Red },
        { "Green", Colors.Green },
        { "Blue", Colors.Blue },
        { "Yellow", Colors.Yellow },
        { "Cyan", Colors.Cyan },
        { "Magenta", Colors.Magenta },
        { "NavyBlue", Colors.NavyBlue },
        { "DarkBlue", Colors.DarkBlue },
        { "DukeBlue", Colors.DukeBlue },
        { "MediumBlue", Colors.MediumBlue },
        { "DarkGreen", Colors.DarkGreen },
        { "BangladeshGreen", Colors.BangladeshGreen },
        { "SeaBlue", Colors.SeaBlue },
        { "MediumPersianBlue", Colors.MediumPersianBlue },
        { "TrueBlue", Colors.TrueBlue },
        { "BrandeisBlue", Colors.BrandeisBlue },
        { "SpanishViridian", Colors.SpanishViridian },
        { "DarkCyan", Colors.DarkCyan },
        { "BlueCola", Colors.BlueCola },
        { "Azure", Colors.Azure },
        { "IslamicGreen", Colors.IslamicGreen },
        { "GoGreen", Colors.GoGreen },
        { "PersianGreen", Colors.PersianGreen },
        { "TiffanyBlue", Colors.TiffanyBlue },
        { "BlueBolt", Colors.BlueBolt },
        { "ElectricGreen", Colors.ElectricGreen },
        { "Malachite", Colors.Malachite },
        { "CaribbeanGreen", Colors.CaribbeanGreen },
        { "DarkTurquoise", Colors.DarkTurquoise },
        { "VividSkyBlue", Colors.VividSkyBlue },
        { "GuppieGreen", Colors.GuppieGreen },
        { "MediumSpringGreen", Colors.MediumSpringGreen },
        { "SeaGreen", Colors.SeaGreen },
        { "Aqua", Colors.Aqua },
        { "BloodRed", Colors.BloodRed },
        { "ImperialPurple", Colors.ImperialPurple },
        { "MetallicViolet", Colors.MetallicViolet },
        { "ChinesePurple", Colors.ChinesePurple },
        { "FrenchViolet", Colors.FrenchViolet },
        { "ElectricIndigo", Colors.ElectricIndigo },
        { "BronzeYellow", Colors.BronzeYellow },
        { "GraniteGray", Colors.GraniteGray },
        { "UclaBlue", Colors.UclaBlue },
        { "Liberty", Colors.Liberty },
        { "SlateBlue", Colors.SlateBlue },
        { "VeryLightBlue", Colors.VeryLightBlue },
        { "Avocado", Colors.Avocado },
        { "RussianGreen", Colors.RussianGreen },
        { "SteelTeal", Colors.SteelTeal },
        { "Rackley", Colors.Rackley },
        { "UnitedNationsBlue", Colors.UnitedNationsBlue },
        { "Blueberry", Colors.Blueberry },
        { "KellyGreen", Colors.KellyGreen },
        { "ForestGreen", Colors.ForestGreen },
        { "PolishedPine", Colors.PolishedPine },
        { "CrystalBlue", Colors.CrystalBlue },
        { "CarolinaBlue", Colors.CarolinaBlue },
        { "BlueJeans", Colors.BlueJeans },
        { "HarlequinGreen", Colors.HarlequinGreen },
        { "Mantis", Colors.Mantis },
        { "VeryLightMalachiteGreen", Colors.VeryLightMalachiteGreen },
        { "MediumAquamarine", Colors.MediumAquamarine },
        { "MediumTurquoise", Colors.MediumTurquoise },
        { "MayaBlue", Colors.MayaBlue },
        { "BrightGreen", Colors.BrightGreen },
        { "ScreaminGreen", Colors.ScreaminGreen },
        { "Aquamarine", Colors.Aquamarine },
        { "ElectricBlue", Colors.ElectricBlue },
        { "DeepRed", Colors.DeepRed },
        { "FrenchPlum", Colors.FrenchPlum },
        { "MardiGras", Colors.MardiGras },
        { "Violet", Colors.Violet },
        { "ElectricViolet", Colors.ElectricViolet },
        { "GambogeOrange", Colors.GambogeOrange },
        { "DeepTaupe", Colors.DeepTaupe },
        { "ChineseViolet", Colors.ChineseViolet },
        { "RoyalPurple", Colors.RoyalPurple },
        { "MediumPurple", Colors.MediumPurple },
        { "VioletsAreBlue", Colors.VioletsAreBlue },
        { "Olive", Colors.Olive },
        { "Shadow", Colors.Shadow },
        { "TaupeGray", Colors.TaupeGray },
        { "CoolGrey", Colors.CoolGrey },
        { "Ube", Colors.Ube },
        { "AppleGreen", Colors.AppleGreen },
        { "Asparagus", Colors.Asparagus },
        { "DarkSeaGreen", Colors.DarkSeaGreen },
        { "PewterBlue", Colors.PewterBlue },
        { "LightCobaltBlue", Colors.LightCobaltBlue },
        { "FrenchSkyBlue", Colors.FrenchSkyBlue },
        { "AlienArmpit", Colors.AlienArmpit },
        { "PastelGreen", Colors.PastelGreen },
        { "PearlAqua", Colors.PearlAqua },
        { "MiddleBlueGreen", Colors.MiddleBlueGreen },
        { "PaleCyan", Colors.PaleCyan },
        { "Chartreuse", Colors.Chartreuse },
        { "LightGreen", Colors.LightGreen },
        { "DarkCandyAppleRed", Colors.DarkCandyAppleRed },
        { "JazzberryJam", Colors.JazzberryJam },
        { "Flirt", Colors.Flirt },
        { "HeliotropeMagenta", Colors.HeliotropeMagenta },
        { "VividMulberry", Colors.VividMulberry },
        { "ElectricPurple", Colors.ElectricPurple },
        { "WindsorTan", Colors.WindsorTan },
        { "ElectricBrown", Colors.ElectricBrown },
        { "TurkishRose", Colors.TurkishRose },
        { "PearlyPurple", Colors.PearlyPurple },
        { "RichLilac", Colors.RichLilac },
        { "LavenderIndigo", Colors.LavenderIndigo },
        { "DarkGoldenrod", Colors.DarkGoldenrod },
        { "Bronze", Colors.Bronze },
        { "EnglishLavender", Colors.EnglishLavender },
        { "OperaMauve", Colors.OperaMauve },
        { "Lavender", Colors.Lavender },
        { "BrightLavender", Colors.BrightLavender },
        { "LightGold", Colors.LightGold },
        { "OliveGreen", Colors.OliveGreen },
        { "Sage", Colors.Sage },
        { "SilverFoil", Colors.SilverFoil },
        { "WildBlueYonder", Colors.WildBlueYonder },
        { "MaximumBluePurple", Colors.MaximumBluePurple },
        { "VividLimeGreen", Colors.VividLimeGreen },
        { "JuneBud", Colors.JuneBud },
        { "YellowGreen", Colors.YellowGreen },
        { "LightMossGreen", Colors.LightMossGreen },
        { "Crystal", Colors.Crystal },
        { "FreshAir", Colors.FreshAir },
        { "SpringBud", Colors.SpringBud },
        { "Inchworm", Colors.Inchworm },
        { "Menthol", Colors.Menthol },
        { "MagicMint", Colors.MagicMint },
        { "Celeste", Colors.Celeste },
        { "RossoCorsa", Colors.RossoCorsa },
        { "RoyalRed", Colors.RoyalRed },
        { "MexicanPink", Colors.MexicanPink },
        { "DeepMagenta", Colors.DeepMagenta },
        { "Phlox", Colors.Phlox },
        { "Tenn", Colors.Tenn },
        { "IndianRed", Colors.IndianRed },
        { "Blush", Colors.Blush },
        { "SuperPink", Colors.SuperPink },
        { "Orchid", Colors.Orchid },
        { "Heliotrope", Colors.Heliotrope },
        { "HarvestGold", Colors.HarvestGold },
        { "RawSienna", Colors.RawSienna },
        { "NewYorkPink", Colors.NewYorkPink },
        { "MiddlePurple", Colors.MiddlePurple },
        { "DeepMauve", Colors.DeepMauve },
        { "BrightLilac", Colors.BrightLilac },
        { "MustardYellow", Colors.MustardYellow },
        { "EarthYellow", Colors.EarthYellow },
        { "Tan", Colors.Tan },
        { "PaleChestnut", Colors.PaleChestnut },
        { "PinkLavender", Colors.PinkLavender },
        { "Mauve", Colors.Mauve },
        { "Citrine", Colors.Citrine },
        { "ChineseGreen", Colors.ChineseGreen },
        { "MediumSpringBud", Colors.MediumSpringBud },
        { "PastelGray", Colors.PastelGray },
        { "LightSilver", Colors.LightSilver },
        { "PaleLavender", Colors.PaleLavender },
        { "MaximumGreenYellow", Colors.MaximumGreenYellow },
        { "Mindaro", Colors.Mindaro },
        { "TeaGreen", Colors.TeaGreen },
        { "Nyanza", Colors.Nyanza },
        { "LightCyan", Colors.LightCyan },
        { "VividRaspberry", Colors.VividRaspberry },
        { "BrightPink", Colors.BrightPink },
        { "FashionFuchsia", Colors.FashionFuchsia },
        { "ShockingPink", Colors.ShockingPink },
        { "Fuchsia", Colors.Fuchsia },
        { "VividOrange", Colors.VividOrange },
        { "PastelRed", Colors.PastelRed },
        { "Strawberry", Colors.Strawberry },
        { "HotPink", Colors.HotPink },
        { "LightDeepPink", Colors.LightDeepPink },
        { "AmericanOrange", Colors.AmericanOrange },
        { "Coral", Colors.Coral },
        { "Tulip", Colors.Tulip },
        { "TickleMePink", Colors.TickleMePink },
        { "PrincessPerfume", Colors.PrincessPerfume },
        { "FuchsiaPink", Colors.FuchsiaPink },
        { "ChineseYellow", Colors.ChineseYellow },
        { "Rajah", Colors.Rajah },
        { "MacaroniAndCheese", Colors.MacaroniAndCheese },
        { "Melon", Colors.Melon },
        { "LavenderPink", Colors.LavenderPink },
        { "RichBrilliantLavender", Colors.RichBrilliantLavender },
        { "Golden", Colors.Golden },
        { "NaplesYellow", Colors.NaplesYellow },
        { "Jasmine", Colors.Jasmine },
        { "LightOrange", Colors.LightOrange },
        { "PalePink", Colors.PalePink },
        { "PinkLace", Colors.PinkLace },
        { "LaserLemon", Colors.LaserLemon },
        { "PastelYellow", Colors.PastelYellow },
        { "Calamansi", Colors.Calamansi },
        { "Cream", Colors.Cream },
        { "VampireBlack", Colors.VampireBlack },
        { "ChineseBlack", Colors.ChineseBlack },
        { "EerieBlack", Colors.EerieBlack },
        { "RaisinBlack", Colors.RaisinBlack },
        { "DarkCharcoal", Colors.DarkCharcoal },
        { "BlackOlive", Colors.BlackOlive },
        { "OuterSpace", Colors.OuterSpace },
        { "DarkLiver", Colors.DarkLiver },
        { "DavysGrey", Colors.DavysGrey },
        { "DimGray", Colors.DimGray },
        { "SonicSilver", Colors.SonicSilver },
        { "CssGray", Colors.CssGray },
        { "PhilippineGray", Colors.PhilippineGray },
        { "SpanishGray", Colors.SpanishGray },
        { "PhilippineSilver", Colors.PhilippineSilver },
        { "Gray", Colors.Gray },
        { "SilverSand", Colors.SilverSand },
        { "AmericanSilver", Colors.AmericanSilver },
        { "Gainsboro", Colors.Gainsboro },
        { "Platinum", Colors.Platinum },
        { "BrightGray", Colors.BrightGray }
    };
}
