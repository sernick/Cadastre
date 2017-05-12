using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Xml;

using ClosedXML.Excel;


namespace Cadastre
{
	internal class Program
	{
		#region Methods

		private static void Main(string[] args)
		{
			string scriptDirectoryPath = AppDomain.CurrentDomain.BaseDirectory;
			if (!Path.IsPathRooted(scriptDirectoryPath))
			{
				return;
			}

			var scriptDirectory = new DirectoryInfo(scriptDirectoryPath);
			DirectoryInfo excerptsDirectory = scriptDirectory.Parent;

			if (excerptsDirectory == null || !excerptsDirectory.Exists)
			{
				return;
			}

			const StringComparison comparison = StringComparison.InvariantCultureIgnoreCase;
			var estates = new List<Estate>();

			foreach (FileInfo file in excerptsDirectory.EnumerateFiles("*.xml"))
			{
				var builder = new StringBuilder();

				string[] lines = File.ReadAllLines(file.FullName);
				foreach (string line in lines)
				{
					if (!line.StartsWith("<?"))
					{
						if (!string.IsNullOrEmpty(line))
						{
							builder.Append(line);
							builder.Append(Environment.NewLine);
						}
					}
				}
				string text = builder.ToString();

				var document = new XmlDocument();
				try
				{
					document.LoadXml(text);
				}
				catch (Exception e)
				{
					Console.WriteLine(file.FullName);
					Console.WriteLine(e);
					continue;
				}

				foreach (XmlNode kpoks in document.ChildNodes)
				{
					if (string.Equals(kpoks.Name, "KPOKS", comparison))
					{
						foreach (XmlNode realty in kpoks.ChildNodes)
						{
							if (string.Equals(realty.Name, "Realty", comparison))
							{
								foreach (XmlNode flat in realty.ChildNodes)
								{
									if (string.Equals(flat.Name, "Flat", comparison))
									{
										var estate = new Estate();
										estates.Add(estate);

										if (flat.Attributes != null)
										{
											foreach (XmlAttribute a in flat.Attributes)
											{
												if (string.Equals(a.Name, "CadastralNumber", comparison))
												{
													estate.CadastralNumber = a.Value;
												}
												else if (string.Equals(a.Name, "DateCreated", comparison))
												{
													estate.DateCreated = a.Value;
												}
											}
										}

										foreach (XmlNode flatChild in flat.ChildNodes)
										{
											if (string.Equals(flatChild.Name, "CadastralBlock", comparison))
											{
												estate.CadastralBlock = flatChild.InnerText;
											}
											else if (string.Equals(flatChild.Name, "ObjectType", comparison))
											{
												estate.ObjectType = flatChild.InnerText;
											}
											else if (string.Equals(flatChild.Name, "Area", comparison))
											{
												estate.Area = flatChild.InnerText;
											}
											else if (string.Equals(flatChild.Name, "Address", comparison))
											{
												foreach (XmlNode addressChild in flatChild.ChildNodes)
												{
													if (string.Equals(addressChild.Name, "adrs:OKATO", comparison))
													{
														estate.Okato = addressChild.InnerText;
													}
													else if (string.Equals(addressChild.Name, "adrs:KLADR", comparison))
													{
														estate.Kladr = addressChild.InnerText;
													}
													else if (string.Equals(addressChild.Name, "adrs:PostalCode", comparison))
													{
														estate.PostalCode = addressChild.InnerText;
													}
													else if (string.Equals(addressChild.Name, "adrs:Region", comparison))
													{
														estate.Region = addressChild.InnerText;
													}
													else if (string.Equals(addressChild.Name, "adrs:UrbanDistrict", comparison))
													{
														string type = null;
														string name = null;

														if (addressChild.Attributes != null)
														{
															foreach (XmlAttribute a in addressChild.Attributes)
															{
																if (string.Equals(a.Name, "Type", comparison))
																{
																	type = a.Value;
																}
																else if (string.Equals(a.Name, "Name", comparison))
																{
																	name = a.Value;
																}
															}
														}

														estate.UrbanDistrict = $"{name} {type}";
													}
													else if (string.Equals(addressChild.Name, "adrs:Street", comparison))
													{
														string type = null;
														string name = null;

														if (addressChild.Attributes != null)
														{
															foreach (XmlAttribute a in addressChild.Attributes)
															{
																if (string.Equals(a.Name, "Type", comparison))
																{
																	type = a.Value;
																}
																else if (string.Equals(a.Name, "Name", comparison))
																{
																	name = a.Value;
																}
															}
														}

														estate.Street = $"{name} {type}";
													}
													else if (string.Equals(addressChild.Name, "adrs:Level1", comparison))
													{
														if (addressChild.Attributes != null)
														{
															foreach (XmlAttribute a in addressChild.Attributes)
															{
																if (string.Equals(a.Name, "Value", comparison))
																{
																	estate.Level1 = a.Value;
																}
															}
														}
													}
													else if (string.Equals(addressChild.Name, "adrs:Apartment", comparison))
													{
														if (addressChild.Attributes != null)
														{
															foreach (XmlAttribute a in addressChild.Attributes)
															{
																if (string.Equals(a.Name, "Value", comparison))
																{
																	estate.Apartment = a.Value;
																}
															}
														}
													}
												}
											}
											else if (string.Equals(flatChild.Name, "CadastralCost", comparison))
											{
												if (flatChild.Attributes != null)
												{
													foreach (XmlAttribute a in flatChild.Attributes)
													{
														if (string.Equals(a.Name, "Value", comparison))
														{
															estate.CadastralCost = a.Value;
														}
														else if (string.Equals(a.Name, "Unit", comparison))
														{
															estate.CadastralCostUnit = a.Value;
														}
													}
												}
											}
											else if (string.Equals(flatChild.Name, "Rights", comparison))
											{
												foreach (XmlNode rightNode in flatChild.ChildNodes)
												{
													if (string.Equals(rightNode.Name, "Right", comparison))
													{
														var right = new Right();
														estate.Rights.Add(right);

														foreach (XmlNode rightChild in rightNode.ChildNodes)
														{
															if (string.Equals(rightChild.Name, "Name", comparison))
															{
																right.Name = rightChild.InnerText;
															}
															else if (string.Equals(rightChild.Name, "Type", comparison))
															{
																right.Type = rightChild.InnerText;
															}
															else if (string.Equals(rightChild.Name, "Owners", comparison))
															{
																foreach (XmlNode ownerNode in rightChild.ChildNodes)
																{
																	if (string.Equals(ownerNode.Name, "Owner", comparison))
																	{
																		var owner = new Owner();
																		right.Owners.Add(owner);

																		foreach (XmlNode person in ownerNode.ChildNodes)
																		{
																			if (string.Equals(person.Name, "Person", comparison))
																			{
																				foreach (XmlNode child in person.ChildNodes)
																				{
																					if (string.Equals(child.Name, "FamilyName", comparison))
																					{
																						owner.FamilyName = child.InnerText;
																					}
																					else if (string.Equals(child.Name, "FirstName", comparison))
																					{
																						owner.FirstName = child.InnerText;
																					}
																					else if (string.Equals(child.Name, "Patronymic", comparison))
																					{
																						owner.Patronymic = child.InnerText;
																					}
																				}
																			}
																		}
																	}
																}
															}
															else if (string.Equals(rightChild.Name, "Registration", comparison))
															{
																foreach (XmlNode child in rightChild.ChildNodes)
																{
																	if (string.Equals(child.Name, "RegNumber", comparison))
																	{
																		right.RegNumber = child.InnerText;
																	}
																	else if (string.Equals(child.Name, "RegDate", comparison))
																	{
																		right.RegDate = child.InnerText;
																	}
																}
															}
														}
													}
												}
											}
										}

										break;
									}
								}
								break;
							}
						}
						break;
					}
				}
			}

			var workbook = new XLWorkbook();
			IXLWorksheet worksheet = workbook.Worksheets.Add("Flats");

			var headers = new[]
						  {
							  "Квартира",
							  "ОКАТО",
							  "КЛАДР",
							  "Индекс",
							  "Регион",
							  "Район",
							  "Улица",
							  "Дом",
							  "Кадастровый номер",
							  "Кадастровый блок",
							  "Тип объекта",
							  "Площадь",
							  "Кадастровая стоимость",
							  "Собственность",
							  "Тип собственности",
							  "Регистрационный номер",
							  "Дата регистрации",
							  "ФИО"
						  };

			char[] alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

			for (int i = 0; i < headers.Length; i++)
			{
				IXLCell a1 = worksheet.Cell($"{alphabet[i]}1");
				a1.Value = headers[i];
				a1.SetDataType(XLCellValues.Text);

				IXLStyle style = a1.Style;
				style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
				style.Font.Bold = true;
			}

			int row = 1;
			foreach (Estate estate in estates)
			{
				row++;

				IXLCell a = worksheet.Cell($"A{row}");
				a.Value = estate.Apartment;
				a.SetDataType(XLCellValues.Text);

				IXLCell b = worksheet.Cell($"B{row}");
				b.Value = estate.Okato;
				b.SetDataType(XLCellValues.Text);

				IXLCell c = worksheet.Cell($"C{row}");
				c.Value = estate.Kladr;
				c.SetDataType(XLCellValues.Text);

				IXLCell d = worksheet.Cell($"D{row}");
				d.Value = estate.PostalCode;
				d.SetDataType(XLCellValues.Text);

				IXLCell e = worksheet.Cell($"E{row}");
				e.Value = estate.Region;
				e.SetDataType(XLCellValues.Text);

				IXLCell f = worksheet.Cell($"F{row}");
				f.Value = estate.UrbanDistrict;
				f.SetDataType(XLCellValues.Text);

				IXLCell g = worksheet.Cell($"G{row}");
				g.Value = estate.Street;
				g.SetDataType(XLCellValues.Text);

				IXLCell h = worksheet.Cell($"H{row}");
				h.Value = estate.Level1;
				h.SetDataType(XLCellValues.Text);

				IXLCell i = worksheet.Cell($"I{row}");
				i.Value = estate.CadastralNumber;
				i.SetDataType(XLCellValues.Text);

				IXLCell j = worksheet.Cell($"J{row}");
				j.Value = estate.CadastralBlock;
				j.SetDataType(XLCellValues.Text);

				IXLCell k = worksheet.Cell($"K{row}");
				k.Value = estate.ObjectType;
				k.SetDataType(XLCellValues.Text);

				IXLCell l = worksheet.Cell($"L{row}");
				l.Value = estate.Area;

				IXLCell m = worksheet.Cell($"M{row}");
				m.Value = estate.CadastralCost;

				foreach (Right right in estate.Rights)
				{
					row++;

					IXLCell n = worksheet.Cell($"N{row}");
					n.Value = right.Name;
					n.SetDataType(XLCellValues.Text);

					IXLCell o = worksheet.Cell($"O{row}");
					o.Value = right.Type;
					o.SetDataType(XLCellValues.Text);

					IXLCell p = worksheet.Cell($"P{row}");
					p.Value = right.RegNumber;
					p.SetDataType(XLCellValues.Text);

					IXLCell q = worksheet.Cell($"Q{row}");
					q.Value = right.RegDate;

					foreach (Owner owner in right.Owners)
					{
						row++;

						IXLCell r = worksheet.Cell($"R{row}");
						r.Value = $"{owner.FamilyName} {owner.FirstName} {owner.Patronymic}";
						r.SetDataType(XLCellValues.Text);
					}
				}

				row++;
			}

			worksheet.Columns(1, 18).AdjustToContents();

			string xlsxPath = Path.Combine(excerptsDirectory.FullName, $"{DateTime.Now:yyyy-MM-dd HH-mm-ss}.xlsx");

			try
			{
				workbook.SaveAs(xlsxPath);
				Process.Start(xlsxPath);
			}
			catch (Exception e)
			{
				Console.WriteLine(e);
			}

			Console.ReadLine();
		}

		#endregion
	}
}