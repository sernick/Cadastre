using System.Collections.Generic;


namespace Cadastre
{
	internal class Estate
	{
		#region Properties

		public List<Right> Rights
		{
			get;
		} = new List<Right>();

		public string Apartment
		{
			get;
			set;
		}

		public string Area
		{
			get;
			set;
		}

		public string CadastralBlock
		{
			get;
			set;
		}

		public string CadastralCost
		{
			get;
			set;
		}

		public string CadastralCostUnit
		{
			get;
			set;
		}

		public string CadastralNumber
		{
			get;
			set;
		}

		public string DateCreated
		{
			get;
			set;
		}

		public string Kladr
		{
			get;
			set;
		}

		public string Level1
		{
			get;
			set;
		}

		public string ObjectType
		{
			get;
			set;
		}

		public string Okato
		{
			get;
			set;
		}

		public string PostalCode
		{
			get;
			set;
		}

		public string Region
		{
			get;
			set;
		}

		public string Street
		{
			get;
			set;
		}

		public string UrbanDistrict
		{
			get;
			set;
		}

		#endregion
	}
}