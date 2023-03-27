using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExaminationTicket1.Modals
{
    public class WindowService
    {
        public WindowService(string title, decimal costSqM)
        {
            Title = title;
            CostSqM = costSqM;
        }

        public string Title { get; }
        public decimal CostSqM { get; }
    }
}
