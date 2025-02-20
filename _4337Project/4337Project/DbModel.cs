using System;

public class RentalRecord
{
    public int Id { get; set; }
    public string OrderCode { get; set; }
    public DateTime CreationDate { get; set; }
    public string OrderTime { get; set; }
    public string ClientCode { get; set; }
    public string Service { get; set; }
    public string Status { get; set; }
    public DateTime? CloseDate { get; set; }
    public TimeSpan RentalTime { get; set; }
}