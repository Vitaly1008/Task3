using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3
{
    internal class Request
    {
        // код заявки
        private int requestCode;
		public int RequestCode
        {
			get { return requestCode; }
			set { requestCode = value; }
		}

        // код товара
        private int productCode;
		public int ProductCode
        {
			get { return productCode; }
			set { productCode = value; }
		}

        // код клиента
        private int clientCode;
		public int ClientCode
        {
			get { return clientCode; }
			set { clientCode = value; }
		}

        //номер заявки
        private int requestNomber;
		public int RequestNomber
        {
			get { return requestNomber; }
			set { requestNomber = value; }
		}

        //требуемое количество
        private int quantity;
		public int Quantity
        {
			get { return quantity; }
			set { quantity = value; }
		}

        //дата размещения
        private DateTime date;
		public DateTime Date
		{
			get { return date; }
			set { date = value; }
		}

		//конструктор
		public Request(int requestCode, int productCode, int clientCode, int requestNomber, int quantity, DateTime date)
        {
            this.requestCode = requestCode;
			this.productCode = productCode;
			this.clientCode = clientCode;
			this.requestNomber = requestNomber;
			this.quantity = quantity;
			this.date = date;
        }
    }
}
