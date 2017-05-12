using System.Collections.Generic;


namespace Cadastre
{
	internal class Right
	{
		#region Properties

		public List<Owner> Owners
		{
			get;
		} = new List<Owner>();

		public string Name
		{
			get;
			set;
		}

		public string RegDate
		{
			get;
			set;
		}

		public string RegNumber
		{
			get;
			set;
		}

		public string Type
		{
			get;
			set;
		}

		#endregion
	}
}