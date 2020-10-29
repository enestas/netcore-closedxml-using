using System;
using System.Collections.Generic;

namespace ClosedXMLExample.Models {

public class Order {

    public Order() {
        OrderItems = new List<OrderItem>();
    }

    public int IDOrder { get; set; }
    public DateTime OrderDate { get; set; }
    public string CustomerName { get; set; }
    public string CustomerAddress { get; set; }
    public string CustomerPhone { get; set; }
    public List<OrderItem> OrderItems { get; set; }
}

public class OrderItem {
    public int OrderNo { get; set; }
    public string ProductBrand { get; set; }
    public string ProductName { get; set; }
    public decimal Price { get; set; }
    public decimal Quantity { get; set; }
}
}
