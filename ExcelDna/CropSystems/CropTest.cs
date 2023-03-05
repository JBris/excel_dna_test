using System.IO;
using System.Linq;
using System.Reflection;
using ExcelDna.Integration;
using Microsoft.Data.Analysis;


/// <summary>
/// Simple test class to see if dataframes can interact with Excel.
/// </summary>
public static class CropTest
{

    /// <summary>
    /// Print crop data into an Excel spreadsheet.
    /// </summary>
    /// <returns>The crop dataframe as a string</returns>
    [ExcelFunction(Description = "Load Dataframe Data into Excel")]
    public static string PrintCropData()
    {
        DataFrame cropData = LoadCropData();
        string cropDataString = cropData.ToString();
        return cropDataString;
    }


    /// <summary>
    /// Load crop data as an embedded resource into a dataframe
    /// </summary>
    /// <returns>A dataframe containing the crop data</returns>
    public static DataFrame LoadCropData()
    {
        var assembly = Assembly.GetExecutingAssembly();
        string fileName = "crop_data.csv";
        string resourceName = assembly.GetManifestResourceNames()
            .Single(str => str.EndsWith(fileName));

        Stream csv = assembly.GetManifestResourceStream(resourceName);
        DataFrame cropData = DataFrame.LoadCsv(csv);
        return cropData;
    }

}