using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Task3
{
    internal class Сlient
    {
		// код клиента
		private int clientCode;
		public int ClientCode
        {
			get { return clientCode; }
			set { clientCode = value; }
		}

		//организация
		private string organization;
		public string Organization
        {
			get { return organization; }
			set { organization = value; }
		}

		//адрес
		private string address;
		public string Address
        {
			get { return address; }
			set { address = value; }
		}

		//контакт
		private string contact;
		public string Contact
        {
			get { return contact; }
			set { contact = value; }
		}

        //конструктор
        public Сlient(int clientCode, string organization, string address, string contact)
        {
            this.clientCode = clientCode;
			this.organization = organization;
			this.address = address;
			this.contact = contact;
        }

    }
}
