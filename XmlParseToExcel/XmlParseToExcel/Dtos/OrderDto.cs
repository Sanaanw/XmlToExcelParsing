using System.Collections.Generic;
using XmlParseToExcel.Dtos;

namespace XmlParseToExcel
{
    public class OrderDto
    {
        public OrderHeaderDto Header { get; set; }
        public OrderPartyDto Parties { get; set; }
        public List<OrderLineDto> Lines { get; set; }
        public int TotalLines { get; set; }
    }
}
