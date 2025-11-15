using FRJ.Tools.SimpleWorkSheet.Components.Sheet;
using FRJ.Tools.SimpleWorkSheet.Components.SimpleCell;

namespace FRJ.Tools.SimpleWorksheetTests;

public class RowHeightTests
{
    [Fact]
    public void SetRowHeight_ExplicitHeight_StoresCorrectly()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetRowHeight(0, 25.0);
        
        Assert.True(sheet.ExplicitRowHeights.ContainsKey(0));
        Assert.True(sheet.ExplicitRowHeights[0].IsT0);
        Assert.Equal(25.0, sheet.ExplicitRowHeights[0].AsT0);
    }

    [Fact]
    public void SetRowHeight_AutoExpand_StoresCorrectly()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetRowHeight(0, RowHeight.AutoExpand);
        
        Assert.True(sheet.ExplicitRowHeights.ContainsKey(0));
        Assert.True(sheet.ExplicitRowHeights[0].IsT1);
        Assert.Equal(RowHeight.AutoExpand, sheet.ExplicitRowHeights[0].AsT1);
    }

    [Fact]
    public void SetRowHeight_MultipleRows_AllStored()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetRowHeight(0, 20.0);
        sheet.SetRowHeight(1, 30.0);
        sheet.SetRowHeight(2, RowHeight.AutoExpand);
        
        Assert.Equal(3, sheet.ExplicitRowHeights.Count);
        Assert.Equal(20.0, sheet.ExplicitRowHeights[0].AsT0);
        Assert.Equal(30.0, sheet.ExplicitRowHeights[1].AsT0);
        Assert.Equal(RowHeight.AutoExpand, sheet.ExplicitRowHeights[2].AsT1);
    }

    [Fact]
    public void SetRowHeight_OverwriteExisting_UpdatesValue()
    {
        var sheet = new WorkSheet("Test");
        
        sheet.SetRowHeight(0, 20.0);
        sheet.SetRowHeight(0, 40.0);
        
        Assert.Single(sheet.ExplicitRowHeights);
        Assert.Equal(40.0, sheet.ExplicitRowHeights[0].AsT0);
    }
}
