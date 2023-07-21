using System;
using System.Collections.Generic;

namespace Task1.Models;

public partial class Brand
{
    public Brand() { }
    public int BrandId { get; set; }

    public string BrandName { get; set; } = null!;

    public virtual ICollection<Product> Products { get; set; } = new List<Product>();

}
