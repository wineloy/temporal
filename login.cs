	public static string GetAuth(string username, string password)
	{
		var plainTextBytes = System.Text.Encoding.UTF8.GetBytes( username + ":" + password);
		return "Basic " + Convert.ToBase64String(plainTextBytes);
	}
