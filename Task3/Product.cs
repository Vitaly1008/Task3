using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3
{
    internal class Product
    {
        // код продукта
        private int productCode;
        public int ProductCode
        {
            get { return productCode; }
            set { productCode = value; }
        }

        // название
        private string name;
        public string Name
        {
            get { return name; }
            set { name = value; }
        }

        //единица измерения
        private MeasureEnum measure;
        public MeasureEnum Measure
        {
            get { return measure; }
            set { measure = value; }
        }

        //цена за единицу
        private float unitPrice;
        public float UnitPrice
        {
            get { return unitPrice; }
            set { unitPrice = value; }
        }

        //конструктор
        public Product(int productCode, string name, MeasureEnum measure, float unitPrice)
        {
            this.productCode = productCode;
            this.name = name;
            this.measure = measure;
            this.unitPrice = unitPrice;
        }
    }
}
