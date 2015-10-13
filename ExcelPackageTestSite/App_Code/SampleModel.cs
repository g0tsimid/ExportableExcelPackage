using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for SampleModel
/// </summary>
public class SampleModel
{
    public decimal BaseFee { get; set; }
    public decimal TaxRate { get; set; }
    public decimal Total
    {
        get
        {
            return BaseFee * TaxRate / 100;
        }
    }
    public DateTime TransactionDate { get; set; }
    public string Description { get; set; }

	public SampleModel()
	{
		//
		// TODO: Add constructor logic here
		//
	}

    /// <summary>
    /// Return some sample data to test the exporter.
    /// </summary>
    /// <returns></returns>
    public static List<SampleModel> GetData()
    {
        List<SampleModel> res = new List<SampleModel>();

        res.Add(new SampleModel()
        {
            BaseFee = 200,
            TaxRate = 5,
            TransactionDate = DateTime.Now,
            Description = "Sample sale 1"
        });

        res.Add(new SampleModel()
        {
            BaseFee = 49.2m,
            TaxRate = 5,
            TransactionDate = DateTime.Now,
            Description = "Sample sale 2"
        });

        res.Add(new SampleModel()
        {
            BaseFee = 19.99m,
            TaxRate = 0,
            TransactionDate = DateTime.Now,
            Description = "Sample sale 3, no tax"
        });

        return res;
    }
}