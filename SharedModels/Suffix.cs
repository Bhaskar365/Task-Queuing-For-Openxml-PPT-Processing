using System;
using System.Collections.Generic;

namespace WebApplicationAPI.Models;

public partial class Suffix
{
    public string? ProjectName { get; set; }

    public string? TestName { get; set; }

    public string? TestNameColor { get; set; }

    public double Positive { get; set; }
    public double Neutral { get; set; }
    public double Negative { get; set; }

    public bool TestNameBold { get; set; }

    public string? AverageColor { get; set; }

    public string? ProjectTemplateType { get; set; }
}
