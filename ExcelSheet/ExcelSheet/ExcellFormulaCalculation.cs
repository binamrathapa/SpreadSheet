using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SpreadSheet
{
    public class ExcellFormulaCalculation
    {
        public double GetCalculatedValue(string operationName, IList<double> values)
        {
            double result = 0;
            switch(operationName.ToUpper())
            {
                case "SUM":
                    foreach(var v in values)
                    {
                        result += v;
                    }

                    break;
                case "AVERAGE":                    
                case "MEAN":
                    foreach (var v in values)
                    {
                        result += v;
                    }
                    result = result / values.Count;
                    break;
                case "MEDIAN":
                    int mean;
                    if (values.Count % 2 == 0)
                    {
                        mean = (values.Count / 2) + 1;
                    }
                    else
                    {
                        mean = (values.Count + 1) / 2;
                    }
                    result = values[mean];

                    break;
                case "MODE":

                    break;


            }
            return result;
        }
    }
}
