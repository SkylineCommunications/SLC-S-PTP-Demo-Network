using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Skyline.AppInstaller;
using Skyline.DataMiner.Automation;
using Skyline.DataMiner.Core.DataMinerSystem.Automation;
using Skyline.DataMiner.Core.DataMinerSystem.Common;
using Skyline.DataMiner.Net.AppPackages;
using Skyline.DataMiner.Net.Messages;

/// <summary>
/// DataMiner Script Class.
/// </summary>
internal class Script
{
	/// <summary>
	/// The script entry point.
	/// </summary>
	/// <param name="engine">Provides access to the Automation engine.</param>
	/// <param name="context">Provides access to the installation context.</param>
	[AutomationEntryPoint(AutomationEntryPointType.Types.InstallAppPackage)]
	public void Install(IEngine engine, AppInstallContext context)
	{
		try
		{
			engine.Timeout = new TimeSpan(0, 10, 0);
			engine.GenerateInformation("Starting installation");
			var installer = new AppInstaller(Engine.SLNetRaw, context);
			installer.InstallDefaultContent();

			////string setupContentPath = installer.GetSetupContentDirectory();

			// Custom installation logic can be added here for each individual install package.
			SetAllNewProtocolVersionsAsProduction(engine, installer);

			var dms = engine.GetDms();
			string activePreferGrandMasterId = string.Empty;
			int numericId = 1;

			// Create Views
			int demoView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo");
			int devicesView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - Devices", demoView);
			int gmView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - GM", devicesView);
			int bCView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - BC", devicesView);
			int tcView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - TC", devicesView);
			int slaveView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - Slave", devicesView);

			// Create Elements
			Dictionary<int, object> defaultConfig = new Dictionary<int, object>
			{
				[10017] = 101, // PTP Domain
				[10001] = 1, // P2P Mode
				[10049] = 128, // Grandmaster Priority 1
				[10043] = 6, // Grandmaster Clock Class
				[10045] = 32, // Grandmaster Clock Accuracy,
				[10047] = 15, // Grandmaster Clock Variance
				[10051] = 1, // Grandmaster Priority 2
				[10021] = 3, // Lock Status
				[10027] = 0, // Offset
				[10029] = 0, // Mean Path Delay
				[10023] = 3, // Mode
				[10063] = 37, // UTC Offset
				[10065] = 1, // UTC Offset Valid
				[10067] = 0, // Leap 59
				[10069] = 0, // Leap 61
				[10071] = 1, // Time Tracing
				[10073] = 1, // Frequency Tracing
			};

			Dictionary<int, object> grandmastersConfig = new Dictionary<int, object>()
			{
				[10018] = 101, // PTP Domain
				[10014] = 128, // Priority 1
				[10008] = 6, // Clock Class
				[10010] = 32, // Clock Accuracy
				[10012] = 1, // Clock Variance
				[10020] = 0, // Slave Only
				[10314] = -2, // Grandmaster Announce Rate
				[10316] = -3, // Grandmaster Sync Rate
				[10318] = 1, // Grandmaster Delay Response Rate
				[10324] = 1, // Grandmaster Clock Timescale
				[10326] = 32, // Grandmaster Clock Source
			};

			Dictionary<int, object> boundaryClocksAndSlavesConfig = new Dictionary<int, object>(defaultConfig)
			{
				[10013] = 255, // Priority 1
				[10007] = 6, // Clock Class
				[10009] = 32, // Clock Accuracy
				[10011] = 1, // Clock Variance
				[10015] = 2, // Priority 2
				[10019] = 1, // Slave Only
				[232] = 1, // Interfaces Configuration Mode (Leaf)
				[120] = 1, // Deploy Interfaces
			};

			Dictionary<int, object> boundaryClocksConfig = new Dictionary<int, object>(boundaryClocksAndSlavesConfig)
			{
				[10023] = 2, // Mode
				[10033] = 1, // Parent Port Number
				[10035] = 1, // Parent Stats
				[10037] = 0, // Parent Clock Variance
				[10039] = 0, // Parent Phase Change Rate
				[232] = 2, // Interfaces Configuration Mode (Spine)
			};

			Dictionary<int, object> obtruckBoundaryClocksConfig = new Dictionary<int, object>(boundaryClocksConfig)
			{
				[10013] = 224, // Priority 1
			};

			Dictionary<int, object> singaporeSlavesPriorityConfig = new Dictionary<int, object>(boundaryClocksAndSlavesConfig)
			{
				[10013] = 224, // Priority 1
			};

			Dictionary<int, object> slaveFromDifferentDomainConfig = new Dictionary<int, object>(boundaryClocksAndSlavesConfig)
			{
				[10017] = 111, // PTP Domain
			};

			List<ElementConfig> gms = new List<ElementConfig> // todo: make a config for the GMs that fits the PTP Edge DVE
			{
				new ElementConfig("NY-GM-01", new Dictionary<int, object>(grandmastersConfig) { }),
				new ElementConfig("TYO-GM-01", new Dictionary<int, object>(grandmastersConfig) { }),
			};

			List<ElementConfig> boundaryClocks = new List<ElementConfig>
			{
				new ElementConfig("LA SW[BLUE]-L01", boundaryClocksConfig),
				new ElementConfig("LA SW[BLUE]-L02", boundaryClocksConfig),
				new ElementConfig("LA SW[BLUE]-L03", boundaryClocksConfig),
				new ElementConfig("LA SW[RED]-L01", boundaryClocksConfig),
				new ElementConfig("LA SW[RED]-L02", boundaryClocksConfig),
				new ElementConfig("LA SW[RED]-L03", boundaryClocksConfig),
				new ElementConfig("NY SW-L01", boundaryClocksConfig),
				new ElementConfig("NY SW-L02", boundaryClocksConfig),
				new ElementConfig("NY SW-L03", boundaryClocksConfig),
				new ElementConfig("NY SW-L04", boundaryClocksConfig),
				new ElementConfig("NY SW-L05", boundaryClocksConfig),
				new ElementConfig("TYO SW[BLUE]-L01", boundaryClocksConfig),
				new ElementConfig("TYO SW[BLUE]-L02", boundaryClocksConfig),
				new ElementConfig("TYO SW[BLUE]-L03", boundaryClocksConfig),
				new ElementConfig("TYO SW[RED]-L01", boundaryClocksConfig),
				new ElementConfig("TYO SW[RED]-L02", boundaryClocksConfig),
				new ElementConfig("TYO SW[RED]-L03", boundaryClocksConfig),
				new ElementConfig("OB SW-L01", obtruckBoundaryClocksConfig),
				new ElementConfig("OB SW-L02", obtruckBoundaryClocksConfig),
			};

			List<ElementConfig> transparrentClocks = new List<ElementConfig>
			{
			};

			List<ElementConfig> slaveClocks = new List<ElementConfig>
			{
				// Probe
				new ElementConfig("NY PTP-Probe", boundaryClocksAndSlavesConfig),

				// Slaves
				new ElementConfig("LA SW[BLUE]-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("LA SW[BLUE]-S02", boundaryClocksAndSlavesConfig),
				new ElementConfig("LA SW[RED]-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("LA SW[RED]-S02", boundaryClocksAndSlavesConfig),
				new ElementConfig("NY SW-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("NY SW-S02", boundaryClocksAndSlavesConfig),
				new ElementConfig("OB SW[BLUE]-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("OB SW[RED]-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("PAR SW[BLUE]-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("TYO SW[BLUE]-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("TYO SW[BLUE]-S02", boundaryClocksAndSlavesConfig),
				new ElementConfig("TYO SW[RED]-S01", boundaryClocksAndSlavesConfig),
				new ElementConfig("TYO SW[RED]-S02", boundaryClocksAndSlavesConfig),
				new ElementConfig("SG SW[BLUE]-S01", singaporeSlavesPriorityConfig),
				new ElementConfig("SG SW[RED]-S01", singaporeSlavesPriorityConfig),
				new ElementConfig("PAR SW[RED]-S01", slaveFromDifferentDomainConfig),
			};

			CreateGMs(dms, gms, gmView);
			CreateClocks(dms, boundaryClocks, bCView);
			CreateClocks(dms, new List<ElementConfig>(), tcView); // todo: create Transparent clocks
			CreateClocks(dms, slaveClocks, slaveView);

			// Configure Elements
			foreach (var clock in gms)
			{
				ApplyGMConfig(engine, clock, ref numericId, ref activePreferGrandMasterId);
				numericId++;
			}

			foreach (var clock in boundaryClocks)
			{
				ApplyConfigurationToElement(engine, clock, numericId, activePreferGrandMasterId);
				numericId++;
			}

			foreach (var clock in transparrentClocks)
			{
				ApplyConfigurationToElement(engine, clock, numericId, activePreferGrandMasterId);
				numericId++;
			}

			foreach (var clock in slaveClocks)
			{
				ApplyConfigurationToElement(engine, clock, numericId, activePreferGrandMasterId);
				numericId++;
			}

			// Create DCF Connections
			List<Connection> connections = new List<Connection>
			{
				new Connection(engine, "NY-GM-01.PTP Card 1", "Ethernet 1/1", "NY SW-L01", "Ethernet 100"),
				new Connection(engine, "NY-GM-01.PTP Card 2", "Ethernet 2/1", "NY SW-L01", "Ethernet 101"),
				new Connection(engine, "NY-GM-01.PTP Card 1", "Ethernet 1/2", "NY SW-L02", "Ethernet 100"),
				new Connection(engine, "NY-GM-01.PTP Card 2", "Ethernet 2/2", "NY SW-L02", "Ethernet 101"),
				new Connection(engine, "NY-GM-01.PTP Card 1", "Ethernet 1/3", "NY SW-L03", "Ethernet 100"),
				new Connection(engine, "NY-GM-01.PTP Card 2", "Ethernet 2/3", "NY SW-L03", "Ethernet 101"),
				new Connection(engine, "NY-GM-01.PTP Card 1", "Ethernet 1/4", "NY SW-L04", "Ethernet 100"),
				new Connection(engine, "NY-GM-01.PTP Card 2", "Ethernet 2/4", "NY SW-L04", "Ethernet 101"),
				new Connection(engine, "NY-GM-01.PTP Card 1", "Ethernet 1/4", "NY SW-L05", "Ethernet 100"), // Todo: fix this, the PTP cards only have 4 interfaces
				new Connection(engine, "NY-GM-01.PTP Card 2", "Ethernet 2/4", "NY SW-L05", "Ethernet 101"), // Todo: fix this, the PTP cards only have 4 interfaces

				new Connection(engine, "TYO-GM-01.PTP Card 1", "Ethernet 1/1", "TYO SW[BLUE]-L01", "Ethernet 100"),
				new Connection(engine, "TYO-GM-01.PTP Card 1", "Ethernet 1/2", "TYO SW[BLUE]-L02", "Ethernet 100"),
				new Connection(engine, "TYO-GM-01.PTP Card 1", "Ethernet 1/3", "TYO SW[BLUE]-L03", "Ethernet 100"),
				new Connection(engine, "TYO-GM-01.PTP Card 2", "Ethernet 2/1", "TYO SW[RED]-L01", "Ethernet 100"),
				new Connection(engine, "TYO-GM-01.PTP Card 2", "Ethernet 2/2", "TYO SW[RED]-L02", "Ethernet 100"),
				new Connection(engine, "TYO-GM-01.PTP Card 2", "Ethernet 2/3", "TYO SW[RED]-L03", "Ethernet 100"),

				new Connection(engine, "TYO SW[BLUE]-L01", "Ethernet 1", "TYO SW[BLUE]-S01", "Ethernet 49"),
				new Connection(engine, "TYO SW[BLUE]-L01", "Ethernet 2", "TYO SW[BLUE]-S02", "Ethernet 49"),
				new Connection(engine, "TYO SW[BLUE]-L02", "Ethernet 1", "TYO SW[BLUE]-S01", "Ethernet 50"),
				new Connection(engine, "TYO SW[BLUE]-L02", "Ethernet 2", "TYO SW[BLUE]-S02", "Ethernet 50"),
				new Connection(engine, "TYO SW[BLUE]-L03", "Ethernet 1", "TYO SW[BLUE]-S01", "Ethernet 48"),
				new Connection(engine, "TYO SW[BLUE]-L03", "Ethernet 2", "TYO SW[BLUE]-S02", "Ethernet 48"),

				new Connection(engine, "TYO SW[RED]-L01", "Ethernet 1", "TYO SW[RED]-S01", "Ethernet 49"),
				new Connection(engine, "TYO SW[RED]-L01", "Ethernet 2", "TYO SW[RED]-S02", "Ethernet 49"),
				new Connection(engine, "TYO SW[RED]-L02", "Ethernet 1", "TYO SW[RED]-S01", "Ethernet 50"),
				new Connection(engine, "TYO SW[RED]-L02", "Ethernet 2", "TYO SW[RED]-S02", "Ethernet 50"),
				new Connection(engine, "TYO SW[RED]-L03", "Ethernet 1", "TYO SW[RED]-S01", "Ethernet 48"),
				new Connection(engine, "TYO SW[RED]-L03", "Ethernet 2", "TYO SW[RED]-S02", "Ethernet 48"),

				new Connection(engine, "LA SW[BLUE]-L01", "Ethernet 1", "LA SW[BLUE]-S01", "Ethernet 49"),
				new Connection(engine, "LA SW[BLUE]-L01", "Ethernet 2", "LA SW[BLUE]-S02", "Ethernet 49"),
				new Connection(engine, "LA SW[BLUE]-L02", "Ethernet 1", "LA SW[BLUE]-S01", "Ethernet 50"),
				new Connection(engine, "LA SW[BLUE]-L02", "Ethernet 2", "LA SW[BLUE]-S02", "Ethernet 50"),
				new Connection(engine, "LA SW[BLUE]-L03", "Ethernet 1", "LA SW[BLUE]-S01", "Ethernet 48"),
				new Connection(engine, "LA SW[BLUE]-L03", "Ethernet 2", "LA SW[BLUE]-S02", "Ethernet 48"),

				new Connection(engine, "LA SW[RED]-L01", "Ethernet 1", "LA SW[RED]-S01", "Ethernet 49"),
				new Connection(engine, "LA SW[RED]-L01", "Ethernet 2", "LA SW[RED]-S02", "Ethernet 49"),
				new Connection(engine, "LA SW[RED]-L02", "Ethernet 1", "LA SW[RED]-S01", "Ethernet 50"),
				new Connection(engine, "LA SW[RED]-L02", "Ethernet 2", "LA SW[RED]-S02", "Ethernet 50"),
				new Connection(engine, "LA SW[RED]-L03", "Ethernet 1", "LA SW[RED]-S01", "Ethernet 48"),
				new Connection(engine, "LA SW[RED]-L03", "Ethernet 2", "LA SW[RED]-S02", "Ethernet 48"),

				new Connection(engine, "NY SW-L01", "Ethernet 1", "NY SW-S01", "Ethernet 1"),
				new Connection(engine, "NY SW-L02", "Ethernet 1", "NY SW-S01", "Ethernet 2"),
				new Connection(engine, "NY SW-L03", "Ethernet 1", "NY SW-S01", "Ethernet 3"),
				new Connection(engine, "NY SW-L04", "Ethernet 1", "NY SW-S01", "Ethernet 4"),
				new Connection(engine, "NY SW-L05", "Ethernet 1", "NY SW-S01", "Ethernet 5"),

				new Connection(engine, "NY SW-L01", "Ethernet 2", "NY SW-S02", "Ethernet 1"),
				new Connection(engine, "NY SW-L02", "Ethernet 2", "NY SW-S02", "Ethernet 2"),
				new Connection(engine, "NY SW-L03", "Ethernet 2", "NY SW-S02", "Ethernet 3"),
				new Connection(engine, "NY SW-L04", "Ethernet 2", "NY SW-S02", "Ethernet 4"),
				new Connection(engine, "NY SW-L05", "Ethernet 2", "NY SW-S02", "Ethernet 5"),

				new Connection(engine, "LA SW[BLUE]-S01", "Ethernet 1", "LA SW[BLUE]-S02", "Ethernet 1"),
				new Connection(engine, "TYO SW[BLUE]-S01", "Ethernet 1", "TYO SW[BLUE]-S02", "Ethernet 1"),
				new Connection(engine, "LA SW[RED]-S01", "Ethernet 1", "LA SW[RED]-S02", "Ethernet 1"),
				new Connection(engine, "TYO SW[RED]-S01", "Ethernet 1", "TYO SW[RED]-S02", "Ethernet 1"),

				new Connection(engine, "LA SW[BLUE]-S01", "Ethernet 2", "TYO SW[BLUE]-S01", "Ethernet 2"),
				new Connection(engine, "LA SW[RED]-S01", "Ethernet 2", "TYO SW[RED]-S01", "Ethernet 2"),
				new Connection(engine, "LA SW[BLUE]-S02", "Ethernet 2", "NY SW-S01", "Ethernet 6"),
				new Connection(engine, "LA SW[RED]-S02", "Ethernet 2", "NY SW-S02", "Ethernet 6"),
				new Connection(engine, "TYO SW[BLUE]-S02", "Ethernet 2", "NY SW-S02", "Ethernet 7"),
				new Connection(engine, "TYO SW[RED]-S02", "Ethernet 2", "NY SW-S01", "Ethernet 7"),
				new Connection(engine, "NY SW-S01", "Ethernet 8", "NY SW-S02", "Ethernet 8"),

				new Connection(engine, "NY PTP-Probe", "Ethernet 1", "NY SW-L03", "Ethernet 3"),

				new Connection(engine, "NY SW-S01", "Ethernet 9", "SG SW[BLUE]-S01", "Ethernet 1"),
				new Connection(engine, "NY SW-S02", "Ethernet 9", "SG SW[RED]-S01", "Ethernet 1"),
				new Connection(engine, "PAR SW[BLUE]-S01", "Ethernet 1", "SG SW[BLUE]-S01", "Ethernet 2"),
				new Connection(engine, "PAR SW[RED]-S01", "Ethernet 1", "SG SW[RED]-S01", "Ethernet 2"),
				new Connection(engine, "SG SW[BLUE]-S01", "Ethernet 3", "SG SW[RED]-S01", "Ethernet 3"),

				new Connection(engine, "LA SW[BLUE]-S01", "Ethernet 3", "OB SW[BLUE]-S01", "Ethernet 1"),
				new Connection(engine, "LA SW[RED]-S01", "Ethernet 3", "OB SW[RED]-S01", "Ethernet 1"),
				new Connection(engine, "OB SW[BLUE]-S01", "Ethernet 2", "OB SW[RED]-S01", "Ethernet 2"),

				new Connection(engine, "OB SW[BLUE]-S01", "Ethernet 3", "OB SW-L01", "Ethernet 1"),
				new Connection(engine, "OB SW[BLUE]-S01", "Ethernet 4", "OB SW-L02", "Ethernet 1"),
				new Connection(engine, "OB SW[RED]-S01", "Ethernet 3", "OB SW-L01", "Ethernet 2"),
				new Connection(engine, "OB SW[RED]-S01", "Ethernet 4", "OB SW-L02", "Ethernet 2"),
			};

			foreach (var connection in connections)
			{
				connection.MakeConnection(engine);
			}
		}
		catch (Exception e)
		{
			engine.ExitFail($"Exception encountered during installation: {e}");
		}
	}

	private static int CreateViewIfNotExists(IDms dms, string demoViewName, int parentViewID = -1)
	{
		if (!dms.ViewExists(demoViewName))
		{
			int newViewID = dms.CreateView(new ViewConfiguration(demoViewName, dms.GetView(parentViewID)));
			int retries = 0;
			while (!dms.ViewExists(newViewID) && retries < 20)
			{
				Thread.Sleep(100);
				retries++;
			}

			return newViewID;
		}
		else
		{
			return dms.GetView(demoViewName).Id;
		}
	}

	private static void SetAllNewProtocolVersionsAsProduction(IEngine engine, AppInstaller installer)
	{
		try
		{
			var catalogEntries = installer.ContentParser.GetProtocolPackagesToInstall();

			foreach (var entry in catalogEntries)
			{
				try
				{
					var version = Path.GetFileNameWithoutExtension(entry.FileInfo.FullName).Substring(entry.Name.Length + 1);
					installer.Protocol.SetAsProduction(entry.Name, version, copyTemplates: false);

					engine.GenerateInformation($"Setting '{entry.Name}' protocol version '{version}' as Production.");
				}
				catch (Exception ex)
				{
					engine.GenerateInformation($"Exception while setting '{entry.Name}' protocol version as Production: {ex}");
				}
			}
		}
		catch (Exception ex)
		{
			engine.GenerateInformation($"Exception while setting new protocol versions as production: {ex}");
		}
	}

	private static void CreateGMs(IDms dms, List<ElementConfig> elements, int viewID)
	{
		// Create elements
		IDmsProtocol protocol = dms.GetProtocol("Generic Edge Chassis", "Production");
		IUdp port = new Udp("127.0.0.1", 161);
		ISnmpV1Connection mySnmpV1Connection = new SnmpV1Connection(port);

		foreach (var element in elements)
		{
			if (!dms.ElementExists(element.ElementName))
			{
				ElementConfiguration configuration = new ElementConfiguration(dms, element.ElementName, protocol, new List<IElementConnection> { mySnmpV1Connection });
				configuration.Views.Add(dms.GetView(viewID));
				dms.GetAgents().FirstOrDefault().CreateElement(configuration);
			}
		}
	}

	private static void CreateClocks(IDms dms, List<ElementConfig> elements, int viewID)
	{
		IDmsProtocol protocol = dms.GetProtocol("Generic Switch", "Production");
		IUdp port = new Udp("127.0.0.1", 161);
		ISnmpV1Connection mySnmpV1Connection = new SnmpV1Connection(port);

		foreach (var element in elements)
		{
			if (!dms.ElementExists(element.ElementName))
			{
				ElementConfiguration configuration = new ElementConfiguration(dms, element.ElementName, protocol, new List<IElementConnection> { mySnmpV1Connection });
				configuration.Views.Add(dms.GetView(viewID));
				dms.GetAgents().FirstOrDefault().CreateElement(configuration);
			}
		}
	}

	private static void CalculatePtpPorts(ElementConfig elementConfig, Element elem)
	{
		// Calculate 'PTP Ports'
		var interfaces = elem.GetTablePrimaryKeys(65060);
		elementConfig.ParametersToSet[10005] = interfaces?.Count() ?? 0;
	}

	private static string GenerateClockId(int numericId)
	{
		IEnumerable<string> iteratorByteArray = BitConverter.GetBytes(numericId)
			.Reverse()
			.Select(value => value.ToString("X2"));

		const int NumberOfIntermediaOctets = 1;
		return $"01:BA:{String.Join(":", iteratorByteArray.Take(NumberOfIntermediaOctets))}:FF:FE:{String.Join(":", iteratorByteArray.Skip(NumberOfIntermediaOctets))}";
	}

	private void ApplyGMConfig(IEngine engine, ElementConfig elementConfig, ref int numericId, ref string activePreferGrandMasterId)
	{
		int retries = 0;
		Element elem;
		do
		{
			if (retries != 0)
			{
				Thread.Sleep(200);
			}

			retries++;
			elem = engine.FindElement(elementConfig.ElementName);
		}
		while (elem == null && retries < 100);

		elem.SetParameterByPrimaryKey(310, "1", 6); // Create PTP card 1
		elem.SetParameterByPrimaryKey(310, "2", 6); // Create PTP card 2

		retries = 0;
		Element[] cards;
		do
		{
			if (retries != 0)
			{
				Thread.Sleep(200);
			}

			retries++;
			cards = engine.FindElementsByName(elem.ElementName + ".PTP Card *");
		}
		while (cards.Length != 2 && retries < 100);

		foreach (var card in cards)
		{
			string clockId = GenerateClockId(numericId);
			if (activePreferGrandMasterId == string.Empty)
			{
				activePreferGrandMasterId = clockId;
			}

			elementConfig.ParametersToSet[10016] = numericId; // Priority 2
			CalculatePtpPorts(elementConfig, elem); // PTP Ports
			numericId++;

			foreach (KeyValuePair<int, object> elementConfigSetParameter in elementConfig.ParametersToSet)
			{
				card.SetParameter(elementConfigSetParameter.Key, elementConfigSetParameter.Value);
			}

			card.SetParameter(10005, clockId); // Clock ID
			card.SetParameter(10303, clockId); // GrandMaster Clock ID

			// todo: assign alarmtemplate
		}
	}

	private void ApplyConfigurationToElement(IEngine engine, ElementConfig elementConfig, int numericId, string activePreferGrandMasterId)
	{
		int retries = 0;
		Element elem;
		do
		{
			if (retries != 0)
			{
				Thread.Sleep(200);
			}

			retries++;
			elem = engine.FindElement(elementConfig.ElementName);
		}
		while (elem == null && retries < 100);

		elementConfig.ParametersToSet[10002] = GenerateClockId(numericId); // Clock ID
		elementConfig.ParametersToSet[10040] = activePreferGrandMasterId; // Grandmaster Clock ID
		CalculatePtpPorts(elementConfig,elem); // PTP Ports

		// todo: Calculate Steps removed
		foreach (KeyValuePair<int, object> elementConfigSetParameter in elementConfig.ParametersToSet)
		{
			elem.SetParameter(elementConfigSetParameter.Key, elementConfigSetParameter.Value);
		}

		// todo: assign alarmtemplate
	}

	private class ElementConfig
	{
		public ElementConfig(string name, Dictionary<int, object> parametersToSet)
		{
			ElementName = name;
			ParametersToSet = parametersToSet;
		}

		public string ElementName { get; set; }

		public Dictionary<int, object> ParametersToSet { get; set; }

		public string ClockID { get; set; }
	}

	private class Connection
	{
		private string sourceConnectionNameBuffer;

		private string destinationConnectionNameBuffer;

		public string SourceElementName { get; set; }

		public int SourceInterfaceID { get; set; }

		public string DestinationElementName { get; set; }

		public int DestinationInterfaceID { get; set; }

		public Element SourceElement { get; set; }

		public Element DestinationElement { get; set; }

		private string SourceInterfaceName { get; set; }

		private string DestinationInterfaceName { get; set; }

		public Connection(IEngine engine, string sourceElement, string sourceInterfaceName, string destinationElement, string destinationInterfaceName)
		{
			this.SourceElementName = sourceElement;
			this.SourceInterfaceName = sourceInterfaceName;
			this.DestinationElementName = destinationElement;
			this.DestinationInterfaceName = destinationInterfaceName;

			this.SourceElement = engine.FindElement(sourceElement);
			if (this.SourceElement == null)
			{
				engine.GenerateInformation("The source element was not found: " + sourceElement);
			}
			else
			{
				this.SourceInterfaceID = GetSourceInterfaceID(engine, this.SourceElement, sourceInterfaceName);
			}

			this.DestinationElement = engine.FindElement(destinationElement);
			if (this.DestinationElement == null)
			{
				engine.GenerateInformation("The destination element was not found: " + destinationElement);
			}
			else
			{
				this.DestinationInterfaceID = GetDestinationInterfaceID(engine, this.DestinationElement, destinationInterfaceName);
			}
		}

		public int GetSourceInterfaceID(IEngine engine, Element source, string name)
		{
			if (SourceInterfaceID == 0)
			{
				string foundName;
				SourceInterfaceID = GetInterfaceID(engine, source, name, out foundName);
				SourceInterfaceName = foundName;
			}

			return SourceInterfaceID;
		}

		public int GetDestinationInterfaceID(IEngine engine, Element destination, string name)
		{
			if (DestinationInterfaceID == 0)
			{
				string foundName;
				DestinationInterfaceID = GetInterfaceID(engine, destination, name, out foundName);
				DestinationInterfaceName = foundName;
			}

			return DestinationInterfaceID;
		}

		public string GetSourceConnectionName()
		{
			if (sourceConnectionNameBuffer == null)
			{
				GetConnectionNames();
			}

			return sourceConnectionNameBuffer;
		}

		public string GetDestinationConnectionName()
		{
			if (destinationConnectionNameBuffer == null)
			{
				GetConnectionNames();
			}

			return destinationConnectionNameBuffer;
		}

		public bool Exists(IEngine engine)
		{
			// The check is currently only done on the source element, this should be enough in most cases
			// The check is currently verry inefficient, this could definately be improved
			string[] keys = SourceElement.GetTablePrimaryKeys(65060);

			foreach (string myKey in keys)
			{
				int sourceInterfaceID = Convert.ToInt32(SourceElement.GetParameter(65062, myKey));
				if (sourceInterfaceID == SourceInterfaceID)
				{
					string destinationElement = Convert.ToString(SourceElement.GetParameter(65089, myKey));
					if (destinationElement == this.DestinationElement.DmaId + "/" + this.DestinationElement.ElementId)
					{
						int destinationID = Convert.ToInt32(SourceElement.GetParameter(65064, myKey));
						if (destinationID == DestinationInterfaceID)
						{
							return true;
						}
					}
				}
			}

			return false;
		}

		public void MakeConnection(IEngine engine)
		{
			try
			{
				if (!Exists(engine))
				{
					EditConnection myEditConnectionMessage = new EditConnection()
					{
						Action = ConnectivityEditAction.Add,
						ConnectionID = 0,
						DataMinerID = SourceElement.DmaId,
						HostingDataMinerID = SourceElement.DmaId, // todo: this could be wrong
						EditBothConnections = true,
						SourceElementID = SourceElement.ElementId,
						SourceInterfaceID = SourceInterfaceID,
						SourceName = GetSourceConnectionName(),
						DestinationElement = DestinationElement.DmaId + "/" + DestinationElement.ElementId,
						DestinationInterfaceID = DestinationInterfaceID,
						DestinationName = GetDestinationConnectionName(),
					};
					engine.SendSLNetMessage(myEditConnectionMessage);
				}
				else
				{
					// engine.GenerateInformation("Connection " + myConnection.GetSourceConnectionName() + " on element " + myConnection.SourceElementName + " already Exists.");
				}
			}
			catch (Exception e)
			{
				engine.GenerateInformation("Creating DCF connection for " + GetSourceConnectionName() + " failed: " + e.Message);
			}
		}

		private int GetInterfaceID(IEngine engine, Element element, string name, out string foundName)
		{
			// The check is currently verry inefficient, this could definately be improved
			Interface[] interfaces = element.GetInterfaces();
			foundName = name;
			int nonPerfactMatchID = 0;
			if (interfaces == null)
			{
				throw new Exception("Interface " + name + "not found on element " + element.ElementName);
			}
			else // There could be multiple matching
			{
				foreach (Interface myInterface in interfaces)
				{
					if (myInterface.Name == name) // Perfect match, return immediately
					{
						return myInterface.InterfaceId;
					}
					else if (myInterface.Name.Contains(name)) // Not a perfect match, save, but don't return yet
					{
						foundName = myInterface.Name;
						nonPerfactMatchID = myInterface.InterfaceId;
					}
				}
			}

			return nonPerfactMatchID;
		}

		private void GetConnectionNames()
		{
			string sourceInterfaceName = Convert.ToString(SourceElement.GetParameter(65093, Convert.ToString(SourceInterfaceID)));
			string destinationInterfaceName = Convert.ToString(DestinationElement.GetParameter(65093, Convert.ToString(DestinationInterfaceID)));
			sourceConnectionNameBuffer = sourceInterfaceName + " -> " + destinationInterfaceName;
			destinationConnectionNameBuffer = destinationInterfaceName + " -> " + sourceInterfaceName;
		}
	}
}