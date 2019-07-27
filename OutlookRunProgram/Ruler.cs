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
		internal class Rule
		{
			internal class Action
			{
				string run;
				string args;

				internal Action()
				{
				}

				internal bool SetRun(string text)
				{
					// validate
					run = text;
					return true;
				}

				internal bool SetArgs(string text)
				{
					// validate
					args = text;
					return true;
				}
			}


			internal class Regex
			{
				enum Scope
				{
					subject,
					body,
					to,
					from,
					cc,
					invalid
				}

				Scope scope;
				string regex;

				internal Regex(string text)
				{
					scope = Scope.invalid;
					if (text.Length < 2)
					{
						regex = String.Empty;
						return;
					}

					switch(text[0])
					{
						case 's':
							scope = Scope.subject;
							break;
						case 'b':
							scope = Scope.body;
							break;
						case 't':
							scope = Scope.to;
							break;
						case 'f':
							scope = Scope.from;
							break;
						case 'c':
							scope = Scope.cc;
							break;
					}

					regex = text.Substring(1, text.Length - 1);
				}

				internal bool Invalid()
				{
					return scope == Scope.invalid;
				}
			}

			List<Regex> regexes = new List<Regex>();
			List<Action> actions = new List<Action>();

			internal bool AddRegex(string text)
			{
				Regex regex = new Regex(text);
				if (regex.Invalid())
				{
					return false;
				}

				regexes.Add(regex);
				return true;
			}

			internal void AddAction(Action action)
			{
				actions.Add(action);
			}

			internal bool Invalid()
			{
				return actions.Count == 0;
			}

			internal Rule()
			{

			}
		}

		List<Rule> rules = new List<Rule>();

		public void ReadRules(string directoryPath)
		{
			foreach (var xmlfile in Directory.EnumerateFiles(directoryPath, "*.xml"))
			{
				XmlDocument doc = new XmlDocument();
				doc.Load(xmlfile);
				// TODO: validate document, dont care about it atm
				foreach (XmlNode rule_node in doc.DocumentElement.SelectNodes("/rule"))
				{
					// <rule>
					Rule rule = new Rule();

					var match = rule_node.SelectSingleNode("match");
					var regexes = match.SelectNodes("regex");
					if (regexes.Count == 0)
					{
						// this matches everything, okaaaaaaay, can be useful for *ding*
					}

					foreach (XmlNode regex in regexes)
					{
						var text = regex.InnerText;
						if (!rule.AddRegex(text))
						{
							// bad regex
							return;
						}
					}


					var actions_group = rule_node.SelectSingleNode("actions");
					var actions = actions_group.SelectNodes("action");

					foreach (XmlNode action in actions)
					{
						var run = action.SelectSingleNode("run");
						// store it
						Rule.Action ruleAction = new Rule.Action();
						if (!ruleAction.SetRun(run.InnerText))
						{
							// bad executable
							return;
						}




						var args = action.SelectSingleNode("args");
						if (args != null)
						{
							if (!ruleAction.SetArgs(args.InnerText))
							{
								// bad args
								return;
							}
						}

						rule.AddAction(ruleAction);
					}

					rules.Add(rule);
				}

			}
		}
	}
}
