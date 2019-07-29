using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.Office.Interop.Outlook;

namespace OutlookRunProgram
{
	internal class Ruler
	{
		internal class Rule
		{
			internal class RegexResults
			{
				private readonly Dictionary<Regex.Scope, List<Match>> results = new Dictionary<Regex.Scope, List<Match>>();

				internal void Append(MatchCollection matchCollection, Regex.Scope scope)
				{
					if (!results.ContainsKey(scope))
					{
						results.Add(scope, new List<Match>(matchCollection.Cast<Match>()));
					}

					results[scope].AddRange(matchCollection.Cast<Match>());
				}

				internal Match Get(Regex.Scope scope, int resultNumber)
				{
					var matches = results[scope];
					if (matches.Count <= resultNumber)
					{
						return null;
					}

					return matches[resultNumber];
				}
			}

			internal class Action
			{
				internal class Arg
				{
					private readonly ArgType type;
					private readonly Regex.Scope scope;
					private readonly string text;
					private readonly int resultNumber;

					internal Arg(string text)
					{
						this.text = text;
						this.type = ArgType.text;
						this.scope = Regex.Scope.invalid;
					}

					public Arg(Regex.Scope scope, int resultNumber)
					{
						this.type = ArgType.regex;
						this.text = null;
						this.scope = scope;
						this.resultNumber = resultNumber;
					}

					internal enum ArgType
					{
						text,
						regex
					}

					internal static Arg Create(string text)
					{
						Rule.Regex.Scope scope = Regex.Scope.invalid;
						// eh, simple if'ology would be sufficient
						if (text.StartsWith("$s"))
						{
							scope = Regex.Scope.subject;
						}
						else if (text.StartsWith("$b"))
						{
							scope = Regex.Scope.body;
						}
						else if (text.StartsWith("$t"))
						{
							scope = Regex.Scope.to;
						}
						else if (text.StartsWith("$f"))
						{
							scope = Regex.Scope.from;
						}
						else if (text.StartsWith("$c"))
						{
							scope = Regex.Scope.cc;
						}
						else
						{
							return new Arg(text);
						}

						if (!int.TryParse(text.Substring(2, text.Length - 2), out int resultNumber))
						{
							return null;
						}

						return new Arg(scope, resultNumber);
					}

					internal string GetValue(RegexResults results)
					{
						if (type == ArgType.text)
						{
							return text;
						}
						else if (type == ArgType.regex)
						{
							Match match = results.Get(scope, resultNumber);
							if (match == null)
							{
								return String.Empty;
							}

							return match.Value;
						}
						return String.Empty;
					}
				}

				private string run = string.Empty;
				private readonly List<Arg> args = new List<Arg>();
				private bool hide;
				private bool minimize;
				private bool shellexecute;

				internal Action()
				{
				}

				internal bool SetRun(string text)
				{
					// validate
					run = text;
					return true;
				}

				internal bool AddArg(string text)
				{
					Arg arg = Arg.Create(text);
					if (arg == null)
						return false;

					args.Add(arg);
					return true;
				}

				internal void Run(ref RegexResults results)
				{
					ProcessStartInfo processStartInfo = new ProcessStartInfo(run);
					if (args.Count != 0)
					{
						StringBuilder sb = new StringBuilder();

						foreach (var arg in args)
						{
							sb.Append('"');
							sb.Append(arg.GetValue(results));
							sb.Append("\" ");
						}

						processStartInfo.Arguments = sb.ToString();
					}

					Globals.ThisAddIn.GetLogger().Log($"Running: {processStartInfo.FileName} {processStartInfo.Arguments}");

					processStartInfo.UseShellExecute = shellexecute;
					processStartInfo.CreateNoWindow = hide;

					if(minimize)
					{
						processStartInfo.WindowStyle = ProcessWindowStyle.Minimized;
					}

					Process.Start(processStartInfo);
				}

				internal void SetHide(bool v)
				{
					hide = v;
				}

				internal void SetShellExecute(bool v)
				{
					shellexecute = v;
				}

				internal void SetMinimize(bool v)
				{
					minimize = v;
				}
			}

			internal class Regex
			{
				internal enum Scope
				{
					subject,
					body,
					to,
					from,
					cc,
					invalid
				}

				private readonly Scope scope;
				private readonly string regex;
				private readonly System.Text.RegularExpressions.Regex realRegex;

				internal Regex(string text)
				{
					scope = Scope.invalid;
					if (text.Length < 2)
					{
						regex = string.Empty;
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
					try
					{
						realRegex = new System.Text.RegularExpressions.Regex(
							regex,
							RegexOptions.Compiled | RegexOptions.ExplicitCapture
							);
					}
					catch(System.Exception)
					{
						realRegex = null;
						scope = Scope.invalid;
					}
				}

				internal bool Invalid()
				{
					return scope == Scope.invalid;
				}

				internal bool Match(MailItem item, ref RegexResults results)
				{
					MatchCollection matchCollection = null;
					switch (scope)
					{
						case Scope.subject:
							matchCollection = realRegex.Matches(item.Subject);
							break;
						case Scope.body:
							matchCollection = realRegex.Matches(item.Body);
							break;
						case Scope.from:
							matchCollection = realRegex.Matches(item.SenderName);
							break;
						case Scope.to:
							matchCollection = realRegex.Matches(item.To);
							break;
						case Scope.cc:
							matchCollection = realRegex.Matches(item.CC);
							break;
						default:
							return false;
					}

					if (matchCollection.Count == 0) return false;

					Globals.ThisAddIn.GetLogger().Log($"{scope.ToString()} - {matchCollection.Count}");
					for (int i = 0; i < matchCollection.Count; ++i)
					{
						Globals.ThisAddIn.GetLogger().Log($"{i}. {matchCollection[i].Value}");
					}

					results.Append(matchCollection, scope);

					return true;
				}
			}

			private readonly bool is_final;
			private readonly List<Regex> regexes = new List<Regex>();
			private readonly List<Action> actions = new List<Action>();

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

			internal Rule(bool is_final)
			{
				this.is_final = is_final;
			}

			internal bool IsFinal()
			{
				return is_final;
			}

			internal bool Apply(MailItem item)
			{
				RegexResults results = new RegexResults();

				// 1. Check if the mail match all regexes
				foreach (var regex in regexes)
				{
					if (!regex.Match(item, ref results))
					{
						return false;
					}
				}

				// 2. Perform actions
				foreach (var action in actions)
				{
					action.Run(ref results);
				}

				return true;
			}
		}

		internal void ClearRules()
		{
			Globals.ThisAddIn.GetLogger().Log($"Clearing..");
			rules.Clear();
		}

		internal bool ApplyRules(MailItem item)
		{
			Globals.ThisAddIn.GetLogger().Log($"{item.SenderName} - {item.Subject}");
			bool anyRule = false;
			foreach(var rule in rules)
			{
				bool appliedRule = rule.Apply(item);
				Globals.ThisAddIn.GetLogger().Log($"Rule[{rule.GetHashCode()}] applied? {appliedRule}");
				anyRule |= appliedRule;
				if (rule.IsFinal() && appliedRule)
				{
					break;
				}
			}

			return anyRule;
		}

		private readonly List<Rule> rules = new List<Rule>();

		public bool ReadRules(string directoryPath)
		{
			if (!Directory.Exists(directoryPath))
			{
				Directory.CreateDirectory(directoryPath);
				return false;
			}

			foreach (var xmlfile in Directory.EnumerateFiles(directoryPath, "*.xml"))
			{
				Globals.ThisAddIn.GetLogger().Log($"Loading rules from: {xmlfile}");
				XmlDocument doc = new XmlDocument();
				try
				{
					doc.Load(xmlfile);
				}
				catch (System.Exception)
				{
					Globals.ThisAddIn.GetLogger().Log($"Malformed xml: {xmlfile}");
				}

				foreach (XmlNode rule_node in doc.DocumentElement.SelectNodes("/rule"))
				{
					// <rule>
					Rule rule = new Rule(rule_node.SelectSingleNode("final") != null);

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
							return false;
						}
					}

					var actions = rule_node.SelectSingleNode("actions");
					foreach (XmlNode action in actions.SelectNodes("action"))
					{
						var run = action.SelectSingleNode("run");
						// store it
						Rule.Action ruleAction = new Rule.Action();
						if (!ruleAction.SetRun(run.InnerText))
						{
							// bad executable
							return false;
						}

						ruleAction.SetHide(action.SelectSingleNode("hide") != null);
						ruleAction.SetShellExecute(action.SelectSingleNode("shellexecute") != null);
						ruleAction.SetMinimize(action.SelectSingleNode("minimize") != null);

						var args_node = action.SelectSingleNode("args");
						if (args_node != null)
						{
							foreach (XmlNode arg in args_node.SelectNodes("arg"))
							{
								if (!ruleAction.AddArg(arg.InnerText))
								{
									// bad args
									return false;
								}
							}
						}

						rule.AddAction(ruleAction);
					}

					rules.Add(rule);
				}
			}
			Globals.ThisAddIn.GetLogger().Log($"Got {rules.Count}");
			return true;
		}
	}
}
