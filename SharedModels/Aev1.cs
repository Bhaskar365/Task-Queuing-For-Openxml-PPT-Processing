using System;
using System.Collections.Generic;

namespace WebApplicationAPI.Models;

public partial class Aev1
{
    public string? ProjectName { get; set; }

    public string? TestName { get; set; }

    public string? TestNameColor { get; set; }

    public bool TestNameBold { get; set; }

    public double? Average { get; set; }

    public string? AverageColor { get; set; }

    public string? ProjectTemplateType {get;set;}
}
