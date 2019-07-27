using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OutlookRunProgram
{
	internal class Ruler
	{
		public void ReadRules(string directoryPath)
		{
			foreach (var xmlfile in Directory.EnumerateFiles(directoryPath, "*.xml"))
			{
				XmlDocument doc = new XmlDocument();
				doc.Load(xmlfile);
				// TODO: validate document, dont care about it atm
				foreach (XmlNode rule in doc.DocumentElement.SelectNodes("/rule"))
				{
					// <rule>


					var match = rule.SelectSingleNode("match");
					var regexes = match.SelectNodes("regex");
					if (regexes.Count == 0)
					{
						// this matches everything, okaaaaaaay, can be useful for *ding*
					}

					foreach (XmlNode regex in regexes)
					{
						var text = regex.InnerText;
						if (text.Length == 0)
						{
							// bad xml
						}

						switch (text[0])
						{
							case 's': // subject
							case 'b': // body
							case 'c': // cc
							case 't': // to
							case 'f': // from
								// append text[1:] to regex list
								break;
							default:
								// bad regex
								return;
						}
					}


					var actions_group = rule.SelectSingleNode("actions");
					var actions = actions_group.SelectNodes("action");

					foreach (XmlNode action in actions)
					{
						var run = action.SelectSingleNode("run");
						// store it


						var args = action.SelectSingleNode("args");
						if (args != null)
						{
							// parse it
						}
					}
				}

			}
		}
	}
}
