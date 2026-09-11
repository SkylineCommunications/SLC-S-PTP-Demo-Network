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
/// DataMiner Automation Script for PTP Demo Network Setup.
/// Creates a realistic multi-site PTP domain with GMs, BCs, TCs, and Slaves.
/// Parent Clock IDs are set hierarchically based on upstream nodes.
/// </summary>
internal class Script
{
	private static class PtpParam
	{
		public const int ClockId = 10002;
		public const int ClockIdWrite = 10003;
		public const int P2PModeWrite = 10001;
		public const int PtpDomain = 10017;
		public const int ParentClockIdWrite = 10031;
		public const int GrandmasterClockIdWrite = 10041;
		public const int GrandmasterClockId = 10303;
		public const int Priority1 = 10013;
		public const int Priority2 = 10015;
		public const int GmPriority1 = 10014;
		public const int GmPriority2 = 10020;
		public const int ClockClass = 10007;
		public const int ClockAccuracy = 10009;
		public const int ClockVariance = 10011;
		public const int GmClockClass = 10008;
		public const int GmClockAccuracy = 10010;
		public const int GmClockVariance = 10012;
		public const int AnnounceRate = 10314;
		public const int SyncRate = 10316;
		public const int DelayResponseRate = 10318;
		public const int SlaveOnly = 10019;
		public const int Mode = 10023;
		public const int InterfaceConfigMode = 232;
		public const int ParentPortNumber = 10033;
		public const int ParentStats = 10035;
		public const int ParentClockVariance = 10037;
		public const int ParentPhaseChangeRate = 10039;
		public const int LockStatus = 10021;
		public const int Offset = 10027;
		public const int MeanPathDelay = 10029;
		public const int StepsRemovedWrite = 10025;
		public const int PtpPorts = 10005;
		public const int UtcOffset = 10063;
		public const int UtcOffsetValid = 10065;
		public const int Leap59 = 10067;
		public const int Leap61 = 10069;
		public const int TimeTracing = 10071;
		public const int FrequencyTracing = 10073;
		public const int GmClockTimescale = 10324;
		public const int GmClockSource = 10326;
		public const int DeployInterfaces = 120;
	}

	// Write PIDs for Generic Edge Chassis PTP Card DVE - different numbering from Generic Switch
	private static class GmCardParam
	{
		public const int ClockId = 10005;
		public const int PtpPorts = 10007;
		public const int LockStatus = 10023;
		public const int FrequencyLocked = 6;
	}

	private const int PrimaryDomain = 101;
	private const int TestDomain = 111;
	private const int PrimaryGmPriority = 128;
	private const int SecondaryGmPriority = 129;
	private const int CoreBcPriority = 224;
	private const int AccessBcPriority = 225;
	private const int EdgeBcPriority = 226;
	private const int ClockClassGm = 6;
	private const int ClockClassBc = 6;
	private const int ClockAccuracyValue = 32;
	private const int ClockVarianceValue = 1;
	private const int GmAnnounceRate = -2;
	private const int GmSyncRate = -3;
	private const int GmDelayResponseRate = -3;
	private const int DefaultSwitchPtpPorts = 144;
	private const int TwoStepMode = 1;
	private const int UtcOffsetUtc = 0;
	private const int UtcOffsetAsia = 9 * 3600;
	private const int UtcOffsetAmerica = -5 * 3600;

	// PTP Hierarchy: maps element name to its upstream parent element name
	private static Dictionary<string, string> hierarchyMap = new Dictionary<string, string>
	{
		// Grandmasters (no parent)
		["NY-GM-01"] = null,
		["TYO-GM-01"] = null,

		// Transparent Clocks (parent = upstream connected switch)
		["NY TC-01"] = "NY SW-L01",
		["LA TC-01"] = "LA SW[BLUE]-L01",
		["TYO TC-01"] = "TYO SW[BLUE]-L01",

		// NYC Core BCs (direct to NY-GM-01)
		["NY SW-L01"] = "NY-GM-01",
		["NY SW-L02"] = "NY-GM-01",
		["NY SW-L03"] = "NY-GM-01",
		["NY SW-L04"] = "NY-GM-01",

		// LA Access BCs (sync to NY core BCs, primary to L01 for consistency)
		["LA SW[BLUE]-L01"] = "NY SW-L01",
		["LA SW[BLUE]-L02"] = "NY SW-L01",
		["LA SW[BLUE]-L03"] = "NY SW-L01",
		["LA SW[RED]-L01"] = "NY SW-L02",
		["LA SW[RED]-L02"] = "NY SW-L02",
		["LA SW[RED]-L03"] = "NY SW-L02",

		// Tokyo Core BCs (direct to TYO-GM-01)
		["TYO SW[BLUE]-L01"] = "TYO-GM-01",
		["TYO SW[BLUE]-L02"] = "TYO-GM-01",
		["TYO SW[BLUE]-L03"] = "TYO-GM-01",
		["TYO SW[RED]-L01"] = "TYO-GM-01",
		["TYO SW[RED]-L02"] = "TYO-GM-01",
		["TYO SW[RED]-L03"] = "TYO-GM-01",

		// On-Truck Edge BCs (parent to LA access BCs for remote sync)
		["OB SW-L01"] = "LA SW[BLUE]-L01",
		["OB SW-L02"] = "LA SW[BLUE]-L02",

		// NYC Probe (parent to nearest BC for monitoring)
		["NY PTP-Probe"] = "NY SW-L03",

		// NYC Slaves (parent to NYC core BCs)
		["NY SW-S01"] = "NY SW-L01",
		["NY SW-S02"] = "NY SW-L02",

		// LA Slaves (parent to LA access BCs)
		["LA SW[BLUE]-S01"] = "LA SW[BLUE]-L01",
		["LA SW[BLUE]-S02"] = "LA SW[BLUE]-L01",
		["LA SW[RED]-S01"] = "LA SW[RED]-L01",
		["LA SW[RED]-S02"] = "LA SW[RED]-L01",

		// Tokyo Slaves (parent to Tokyo core BCs)
		["TYO SW[BLUE]-S01"] = "TYO SW[BLUE]-L01",
		["TYO SW[BLUE]-S02"] = "TYO SW[BLUE]-L01",
		["TYO SW[RED]-S01"] = "TYO SW[RED]-L01",
		["TYO SW[RED]-S02"] = "TYO SW[RED]-L01",

		// On-Truck Slaves (parent to on-truck BCs)
		["OB SW[BLUE]-S01"] = "OB SW-L01",
		["OB SW[RED]-S01"] = "OB SW-L02",

		// Singapore/Cross-Region Slaves (parent to NYC slaves for backbone connectivity)
		["SG SW[BLUE]-S01"] = "NY SW-S01",
		["SG SW[RED]-S01"] = "NY SW-S02",

		// Paris Domain 101 (parent to Singapore slaves)
		["PAR SW[BLUE]-S01"] = "SG SW[BLUE]-S01",

		// Paris Domain 111 (isolated, parent to SG slave in same domain if available, else GM)
		["PAR SW[RED]-S01"] = "SG SW[RED]-S01",
	};

	[AutomationEntryPoint(AutomationEntryPointType.Types.InstallAppPackage)]
	public void Install(IEngine engine, AppInstallContext context)
	{
		try
		{
			engine.Timeout = new TimeSpan(0, 15, 0);
			engine.GenerateInformation("=== Starting PTP Demo Network Installation ===");
			var installer = new AppInstaller(Engine.SLNetRaw, context);
			installer.InstallDefaultContent();
			SetAllNewProtocolVersionsAsProduction(engine, installer);

			var dms = engine.GetDms();
			int numericIdCounter = 1;

			// Clock ID mapping: element name -> Clock ID (populated during configuration)
			Dictionary<string, string> clockIdMap = new Dictionary<string, string>();

			int demoView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo");
			int devicesView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - Devices", demoView);
			int gmView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - GM", devicesView);
			int bCView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - BC", devicesView);
			int tcView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - TC", devicesView);
			int slaveView = CreateViewIfNotExists(dms, "DataMiner PTP - Demo - Slave", devicesView);

			var gms = CreateGrandmasterConfigs();
			var boundaryClocks = CreateBoundaryClockConfigs();
			var transparentClocks = CreateTransparentClockConfigs();
			var slaveClocks = CreateSlaveClockConfigs();

			engine.GenerateInformation("Creating Grandmaster elements...");
			CreateGMs(dms, gms, gmView);
			CreateClocks(dms, boundaryClocks, bCView);
			CreateClocks(dms, transparentClocks, tcView);
			CreateClocks(dms, slaveClocks, slaveView);

			engine.GenerateInformation("Configuring Grandmaster clocks...");
			foreach (var gm in gms)
			{
				ApplyGrandmasterConfig(engine, gm, ref numericIdCounter, clockIdMap);
				numericIdCounter++;
			}

			engine.GenerateInformation("Configuring Boundary Clocks...");
			foreach (var bc in boundaryClocks)
			{
				ApplyClockConfiguration(engine, bc, numericIdCounter, clockIdMap);
				numericIdCounter++;
			}

			engine.GenerateInformation("Configuring Transparent Clocks...");
			foreach (var tc in transparentClocks)
			{
				ApplyClockConfiguration(engine, tc, numericIdCounter, clockIdMap);
				numericIdCounter++;
			}

			engine.GenerateInformation("Configuring Slave clocks...");
			foreach (var slave in slaveClocks)
			{
				ApplyClockConfiguration(engine, slave, numericIdCounter, clockIdMap);
				numericIdCounter++;
			}

			engine.GenerateInformation("Creating DCF network topology...");
			var connections = BuildNetworkTopology();
			foreach (var connection in connections)
			{
				connection.MakeConnection(engine);
			}

			engine.GenerateInformation("=== PTP Demo Network Installation Complete ===");
		}
		catch (Exception e)
		{
			engine.ExitFail($"Exception encountered during installation: {e}");
		}
	}

	private List<ElementConfig> CreateGrandmasterConfigs()
	{
		var gmBaseConfig = new Dictionary<int, object>
		{
			[PtpParam.PtpDomain] = PrimaryDomain,
			[PtpParam.GmPriority1] = 128,
			[PtpParam.GmPriority2] = 1,
			[PtpParam.GmClockClass] = ClockClassGm,
			[PtpParam.GmClockAccuracy] = ClockAccuracyValue,
			[PtpParam.GmClockVariance] = ClockVarianceValue,
			[PtpParam.SlaveOnly] = 0,
			[PtpParam.AnnounceRate] = GmAnnounceRate,
			[PtpParam.SyncRate] = GmSyncRate,
			[PtpParam.DelayResponseRate] = GmDelayResponseRate,
			[PtpParam.GmClockTimescale] = 1,
			[PtpParam.GmClockSource] = 32,
		};

		var nyGmConfig = new Dictionary<int, object>(gmBaseConfig)
		{
			[PtpParam.GmPriority1] = PrimaryGmPriority,
		};

		var toyoGmConfig = new Dictionary<int, object>(gmBaseConfig)
		{
			[PtpParam.GmPriority1] = SecondaryGmPriority,
		};

		return new List<ElementConfig>
{
new ElementConfig("NY-GM-01", nyGmConfig),
new ElementConfig("TYO-GM-01", toyoGmConfig),
};
	}

	private List<ElementConfig> CreateBoundaryClockConfigs()
	{
		var bcBaseConfig = new Dictionary<int, object>
		{
			[PtpParam.PtpDomain] = PrimaryDomain,
			[PtpParam.Mode] = 2,
			[PtpParam.SlaveOnly] = 0,
			[PtpParam.InterfaceConfigMode] = 2,
			[PtpParam.ClockClass] = ClockClassBc,
			[PtpParam.ClockAccuracy] = ClockAccuracyValue,
			[PtpParam.ClockVariance] = ClockVarianceValue,
			[PtpParam.Priority2] = 2,
			[PtpParam.ParentPortNumber] = 1,
			[PtpParam.ParentStats] = 1,
			[PtpParam.LockStatus] = 3,
			[PtpParam.Offset] = 0,
			[PtpParam.MeanPathDelay] = 0,
			[PtpParam.DeployInterfaces] = 1,
			[PtpParam.TimeTracing] = 1,
			[PtpParam.FrequencyTracing] = 1,
		};

		var coreBcConfig = new Dictionary<int, object>(bcBaseConfig)
		{
			[PtpParam.Priority1] = CoreBcPriority,
		};

		var accessBcConfig = new Dictionary<int, object>(bcBaseConfig)
		{
			[PtpParam.Priority1] = AccessBcPriority,
		};

		var edgeBcConfig = new Dictionary<int, object>(bcBaseConfig)
		{
			[PtpParam.Priority1] = EdgeBcPriority,
		};

		var bcs = new List<ElementConfig>();

		foreach (var sw in new[] { "NY SW-L01", "NY SW-L02", "NY SW-L03", "NY SW-L04" })
		{
			var config = new Dictionary<int, object>(coreBcConfig)
			{
				[PtpParam.UtcOffset] = UtcOffsetAmerica,
				[PtpParam.UtcOffsetValid] = 1,
			};
			bcs.Add(new ElementConfig(sw, config));
		}

		foreach (var color in new[] { "BLUE", "RED" })
		{
			foreach (var num in new[] { "L01", "L02", "L03" })
			{
				var config = new Dictionary<int, object>(accessBcConfig)
				{
					[PtpParam.UtcOffset] = UtcOffsetAmerica,
					[PtpParam.UtcOffsetValid] = 1,
				};
				bcs.Add(new ElementConfig($"LA SW[{color}]-{num}", config));
			}
		}

		foreach (var color in new[] { "BLUE", "RED" })
		{
			foreach (var num in new[] { "L01", "L02", "L03" })
			{
				var config = new Dictionary<int, object>(coreBcConfig)
				{
					[PtpParam.UtcOffset] = UtcOffsetAsia,
					[PtpParam.UtcOffsetValid] = 1,
				};
				bcs.Add(new ElementConfig($"TYO SW[{color}]-{num}", config));
			}
		}

		foreach (var num in new[] { "L01", "L02" })
		{
			var config = new Dictionary<int, object>(edgeBcConfig)
			{
				[PtpParam.UtcOffset] = UtcOffsetAmerica,
				[PtpParam.UtcOffsetValid] = 1,
			};
			bcs.Add(new ElementConfig($"OB SW-{num}", config));
		}

		return bcs;
	}

	private List<ElementConfig> CreateTransparentClockConfigs()
	{
		var tcBaseConfig = new Dictionary<int, object>
		{
			[PtpParam.PtpDomain] = PrimaryDomain,
			[PtpParam.Mode] = 1,
			[PtpParam.SlaveOnly] = 0,
			[PtpParam.ClockClass] = ClockClassBc,
			[PtpParam.ClockAccuracy] = ClockAccuracyValue,
			[PtpParam.ClockVariance] = ClockVarianceValue,
			[PtpParam.LockStatus] = 3,
			[PtpParam.DeployInterfaces] = 1,
			[PtpParam.TimeTracing] = 1,
			[PtpParam.FrequencyTracing] = 1,
		};

		return new List<ElementConfig>
{
new ElementConfig("NY TC-01", new Dictionary<int, object>(tcBaseConfig)
{
[PtpParam.UtcOffset] = UtcOffsetAmerica,
[PtpParam.UtcOffsetValid] = 1,
}),
new ElementConfig("LA TC-01", new Dictionary<int, object>(tcBaseConfig)
{
[PtpParam.UtcOffset] = UtcOffsetAmerica,
[PtpParam.UtcOffsetValid] = 1,
}),
new ElementConfig("TYO TC-01", new Dictionary<int, object>(tcBaseConfig)
{
[PtpParam.UtcOffset] = UtcOffsetAsia,
[PtpParam.UtcOffsetValid] = 1,
}),
};
	}

	private List<ElementConfig> CreateSlaveClockConfigs()
	{
		var slaveBaseConfig = new Dictionary<int, object>
		{
			[PtpParam.PtpDomain] = PrimaryDomain,
			[PtpParam.Mode] = 3,
			[PtpParam.SlaveOnly] = 0,
			[PtpParam.InterfaceConfigMode] = 1,
			[PtpParam.ClockClass] = ClockClassBc,
			[PtpParam.ClockAccuracy] = ClockAccuracyValue,
			[PtpParam.ClockVariance] = ClockVarianceValue,
			[PtpParam.Priority1] = 255,
			[PtpParam.Priority2] = 2,
			[PtpParam.LockStatus] = 3,
			[PtpParam.Offset] = 0,
			[PtpParam.MeanPathDelay] = 0,
			[PtpParam.DeployInterfaces] = 1,
			[PtpParam.TimeTracing] = 1,
			[PtpParam.FrequencyTracing] = 1,
		};

		var slaves = new List<ElementConfig>();

		slaves.Add(new ElementConfig("NY PTP-Probe", new Dictionary<int, object>(slaveBaseConfig)
		{
			[PtpParam.UtcOffset] = UtcOffsetAmerica,
			[PtpParam.UtcOffsetValid] = 1,
		}));

		slaves.Add(new ElementConfig("NY SW-S01", new Dictionary<int, object>(slaveBaseConfig)
		{
			[PtpParam.UtcOffset] = UtcOffsetAmerica,
			[PtpParam.UtcOffsetValid] = 1,
		}));
		slaves.Add(new ElementConfig("NY SW-S02", new Dictionary<int, object>(slaveBaseConfig)
		{
			[PtpParam.UtcOffset] = UtcOffsetAmerica,
			[PtpParam.UtcOffsetValid] = 1,
		}));

		foreach (var color in new[] { "BLUE", "RED" })
		{
			foreach (var num in new[] { "S01", "S02" })
			{
				slaves.Add(new ElementConfig($"LA SW[{color}]-{num}", new Dictionary<int, object>(slaveBaseConfig)
				{
					[PtpParam.UtcOffset] = UtcOffsetAmerica,
					[PtpParam.UtcOffsetValid] = 1,
				}));
			}
		}

		foreach (var color in new[] { "BLUE", "RED" })
		{
			foreach (var num in new[] { "S01", "S02" })
			{
				slaves.Add(new ElementConfig($"TYO SW[{color}]-{num}", new Dictionary<int, object>(slaveBaseConfig)
				{
					[PtpParam.UtcOffset] = UtcOffsetAsia,
					[PtpParam.UtcOffsetValid] = 1,
				}));
			}
		}

		slaves.Add(new ElementConfig("OB SW[BLUE]-S01", new Dictionary<int, object>(slaveBaseConfig)
		{
			[PtpParam.Priority1] = EdgeBcPriority,
			[PtpParam.UtcOffset] = UtcOffsetAmerica,
			[PtpParam.UtcOffsetValid] = 1,
		}));
		slaves.Add(new ElementConfig("OB SW[RED]-S01", new Dictionary<int, object>(slaveBaseConfig)
		{
			[PtpParam.Priority1] = EdgeBcPriority,
			[PtpParam.UtcOffset] = UtcOffsetAmerica,
			[PtpParam.UtcOffsetValid] = 1,
		}));

		foreach (var color in new[] { "BLUE", "RED" })
		{
			slaves.Add(new ElementConfig($"SG SW[{color}]-S01", new Dictionary<int, object>(slaveBaseConfig)
			{
				[PtpParam.Priority1] = 224,
				[PtpParam.UtcOffset] = UtcOffsetAsia,
				[PtpParam.UtcOffsetValid] = 1,
			}));
		}

		slaves.Add(new ElementConfig("PAR SW[BLUE]-S01", new Dictionary<int, object>(slaveBaseConfig)
		{
			[PtpParam.UtcOffset] = UtcOffsetUtc,
			[PtpParam.UtcOffsetValid] = 1,
		}));

		slaves.Add(new ElementConfig("PAR SW[RED]-S01", new Dictionary<int, object>(slaveBaseConfig)
		{
			[PtpParam.PtpDomain] = TestDomain,
			[PtpParam.UtcOffset] = UtcOffsetUtc,
			[PtpParam.UtcOffsetValid] = 1,
		}));

		return slaves;
	}

	private static void CreateGMs(IDms dms, List<ElementConfig> elements, int viewID)
	{
		IDmsProtocol protocol = dms.GetProtocol("Generic Edge Chassis", "Production");
		IUdp port = new Udp("127.0.0.1", 161);
		ISnmpV1Connection snmpConnection = new SnmpV1Connection(port);

		foreach (var element in elements)
		{
			if (!dms.ElementExists(element.ElementName))
			{
				ElementConfiguration configuration = new ElementConfiguration(dms, element.ElementName, protocol, new List<IElementConnection> { snmpConnection });
				configuration.Views.Add(dms.GetView(viewID));
				dms.GetAgents().FirstOrDefault().CreateElement(configuration);
			}
		}
	}

	private static void CreateClocks(IDms dms, List<ElementConfig> elements, int viewID)
	{
		IDmsProtocol protocol = dms.GetProtocol("Generic Switch", "Production");
		IUdp port = new Udp("127.0.0.1", 161);
		ISnmpV1Connection snmpConnection = new SnmpV1Connection(port);

		foreach (var element in elements)
		{
			if (!dms.ElementExists(element.ElementName))
			{
				ElementConfiguration configuration = new ElementConfiguration(dms, element.ElementName, protocol, new List<IElementConnection> { snmpConnection });
				configuration.Views.Add(dms.GetView(viewID));
				dms.GetAgents().FirstOrDefault().CreateElement(configuration);
			}
		}
	}

	private void ApplyGrandmasterConfig(IEngine engine, ElementConfig elementConfig, ref int numericId, Dictionary<string, string> clockIdMap)
	{
		Element elem = WaitForElement(engine, elementConfig.ElementName);
		elem.SetParameterByPrimaryKey(310, "1", 6);
		elem.SetParameterByPrimaryKey(310, "2", 6);

		Element[] cards = WaitForCards(engine, elem.ElementName);

		foreach (var card in cards)
		{
			string clockId = GenerateClockId(numericId);
			clockIdMap[elementConfig.ElementName] = clockId;

			elementConfig.ParametersToSet[PtpParam.Priority2] = numericId;

			foreach (var param in elementConfig.ParametersToSet)
				card.SetParameter(param.Key, param.Value);

			card.SetParameter(GmCardParam.ClockId, clockId);
			card.SetParameter(GmCardParam.PtpPorts, 4);
			card.SetParameter(GmCardParam.LockStatus, GmCardParam.FrequencyLocked);
			card.SetParameter(PtpParam.GrandmasterClockId, clockId);

			numericId++;
		}
	}

	private void ApplyClockConfiguration(IEngine engine, ElementConfig elementConfig, int numericId, Dictionary<string, string> clockIdMap)
	{
		Element elem = WaitForElement(engine, elementConfig.ElementName);

		string clockId = GenerateClockId(numericId);
		clockIdMap[elementConfig.ElementName] = clockId;
		elementConfig.ParametersToSet[PtpParam.ClockIdWrite] = clockId;

		string parentClockId = string.Empty;
		if (hierarchyMap.ContainsKey(elementConfig.ElementName) && hierarchyMap[elementConfig.ElementName] != null)
		{
			string upstreamNodeName = hierarchyMap[elementConfig.ElementName];
			if (clockIdMap.ContainsKey(upstreamNodeName))
			{
				parentClockId = clockIdMap[upstreamNodeName];
				engine.GenerateInformation($"Setting {elementConfig.ElementName} parent Clock ID to {upstreamNodeName}: {parentClockId}");
			}
			else
			{
				engine.GenerateInformation($"Warning: Parent node '{upstreamNodeName}' for '{elementConfig.ElementName}' not yet configured; will use empty parent");
			}
		}

		string grandmasterClockId = ResolveGrandmasterClockId(elementConfig.ElementName, clockIdMap);
		int stepsRemoved = ResolveStepsRemoved(elementConfig.ElementName);
		elementConfig.ParametersToSet[PtpParam.P2PModeWrite] = TwoStepMode;
		elementConfig.ParametersToSet[PtpParam.ParentClockIdWrite] = parentClockId;
		elementConfig.ParametersToSet[PtpParam.GrandmasterClockIdWrite] = grandmasterClockId;
		elementConfig.ParametersToSet[PtpParam.StepsRemovedWrite] = stepsRemoved;
		elementConfig.ParametersToSet[PtpParam.PtpPorts] = DefaultSwitchPtpPorts;

		foreach (var param in elementConfig.ParametersToSet)
			elem.SetParameter(param.Key, param.Value);
	}

	private static string ResolveGrandmasterClockId(string elementName, Dictionary<string, string> clockIdMap)
	{
		string current = elementName;
		int depth = 0;
		while (depth < 50 && hierarchyMap.ContainsKey(current) && hierarchyMap[current] != null)
		{
			current = hierarchyMap[current];
			depth++;
		}

		if (clockIdMap.ContainsKey(current))
			return clockIdMap[current];

		return string.Empty;
	}

	private static int ResolveStepsRemoved(string elementName)
	{
		string current = elementName;
		int depth = 0;
		while (depth < 50 && hierarchyMap.ContainsKey(current) && hierarchyMap[current] != null)
		{
			current = hierarchyMap[current];
			depth++;
		}

		return depth;
	}

	private static Element WaitForElement(IEngine engine, string elementName, int maxRetries = 100)
	{
		int retries = 0;
		Element elem;
		do
		{
			if (retries != 0) Thread.Sleep(200);
			retries++;
			elem = engine.FindElement(elementName);
		}
		while (elem == null && retries < maxRetries);

		if (elem == null)
			throw new Exception($"Element '{elementName}' did not become available");

		return elem;
	}

	private static Element[] WaitForCards(IEngine engine, string parentElementName, int maxRetries = 100)
	{
		int retries = 0;
		Element[] cards;
		do
		{
			if (retries != 0) Thread.Sleep(200);
			retries++;
			cards = engine.FindElementsByName(parentElementName + ".PTP Card *");
		}
		while (cards.Length != 2 && retries < maxRetries);

		foreach (var card in cards)
			WaitForElementActive(engine, card.ElementName);

		return cards;
	}

	// DVEs are present in FindElements before their parameters are initialised; wait for Active state.
	private static void WaitForElementActive(IEngine engine, string elementName, int maxRetries = 150)
	{
		var dms = engine.GetDms();
		int retries = 0;
		while (retries < maxRetries)
		{
			try
			{
				var elem = engine.FindElement(elementName);
				if (elem != null)
				{
					var dmsElement = dms.GetElement(new Skyline.DataMiner.Core.DataMinerSystem.Common.DmsElementId(elem.DmaId, elem.ElementId));
					if (dmsElement.State == Skyline.DataMiner.Core.DataMinerSystem.Common.ElementState.Active)
						return;
				}
			}
			catch { }
			Thread.Sleep(200);
			retries++;
		}
	}

	private List<Connection> BuildNetworkTopology()
	{
		var connections = new List<Connection>();

		connections.Add(new Connection("NY-GM-01.PTP Card 1", "Ethernet 1/1", "NY SW-L01", "Ethernet 100"));
		connections.Add(new Connection("NY-GM-01.PTP Card 1", "Ethernet 1/2", "NY SW-L02", "Ethernet 100"));
		connections.Add(new Connection("NY-GM-01.PTP Card 1", "Ethernet 1/3", "NY SW-L03", "Ethernet 100"));
		connections.Add(new Connection("NY-GM-01.PTP Card 1", "Ethernet 1/4", "NY SW-L04", "Ethernet 100"));

		connections.Add(new Connection("NY-GM-01.PTP Card 2", "Ethernet 2/1", "NY SW-L01", "Ethernet 101"));
		connections.Add(new Connection("NY-GM-01.PTP Card 2", "Ethernet 2/2", "NY SW-L02", "Ethernet 101"));
		connections.Add(new Connection("NY-GM-01.PTP Card 2", "Ethernet 2/3", "NY SW-L03", "Ethernet 101"));
		connections.Add(new Connection("NY-GM-01.PTP Card 2", "Ethernet 2/4", "NY SW-L04", "Ethernet 101"));

		connections.Add(new Connection("TYO-GM-01.PTP Card 1", "Ethernet 1/1", "TYO SW[BLUE]-L01", "Ethernet 100"));
		connections.Add(new Connection("TYO-GM-01.PTP Card 1", "Ethernet 1/2", "TYO SW[BLUE]-L02", "Ethernet 100"));
		connections.Add(new Connection("TYO-GM-01.PTP Card 1", "Ethernet 1/3", "TYO SW[BLUE]-L03", "Ethernet 100"));

		connections.Add(new Connection("TYO-GM-01.PTP Card 2", "Ethernet 2/1", "TYO SW[RED]-L01", "Ethernet 100"));
		connections.Add(new Connection("TYO-GM-01.PTP Card 2", "Ethernet 2/2", "TYO SW[RED]-L02", "Ethernet 100"));
		connections.Add(new Connection("TYO-GM-01.PTP Card 2", "Ethernet 2/3", "TYO SW[RED]-L03", "Ethernet 100"));

		connections.Add(new Connection("NY PTP-Probe", "Ethernet 1", "NY SW-L03", "Ethernet 3"));

		// Transparent clocks - connected into each regional path.
		connections.Add(new Connection("NY TC-01", "Ethernet 1", "NY SW-L01", "Ethernet 3"));
		connections.Add(new Connection("NY TC-01", "Ethernet 2", "NY SW-L02", "Ethernet 3"));
		connections.Add(new Connection("LA TC-01", "Ethernet 1", "LA SW[BLUE]-L01", "Ethernet 3"));
		connections.Add(new Connection("LA TC-01", "Ethernet 2", "LA SW[RED]-L01", "Ethernet 3"));
		connections.Add(new Connection("TYO TC-01", "Ethernet 1", "TYO SW[BLUE]-L01", "Ethernet 3"));
		connections.Add(new Connection("TYO TC-01", "Ethernet 2", "TYO SW[RED]-L01", "Ethernet 3"));

		foreach (var color in new[] { "BLUE", "RED" })
		{
			connections.Add(new Connection($"LA SW[{color}]-L01", "Ethernet 1", $"LA SW[{color}]-S01", "Ethernet 49"));
			connections.Add(new Connection($"LA SW[{color}]-L01", "Ethernet 2", $"LA SW[{color}]-S02", "Ethernet 49"));
			connections.Add(new Connection($"LA SW[{color}]-L02", "Ethernet 1", $"LA SW[{color}]-S01", "Ethernet 50"));
			connections.Add(new Connection($"LA SW[{color}]-L02", "Ethernet 2", $"LA SW[{color}]-S02", "Ethernet 50"));
			connections.Add(new Connection($"LA SW[{color}]-L03", "Ethernet 1", $"LA SW[{color}]-S01", "Ethernet 48"));
			connections.Add(new Connection($"LA SW[{color}]-L03", "Ethernet 2", $"LA SW[{color}]-S02", "Ethernet 48"));
		}

		foreach (var color in new[] { "BLUE", "RED" })
		{
			connections.Add(new Connection($"TYO SW[{color}]-L01", "Ethernet 1", $"TYO SW[{color}]-S01", "Ethernet 49"));
			connections.Add(new Connection($"TYO SW[{color}]-L01", "Ethernet 2", $"TYO SW[{color}]-S02", "Ethernet 49"));
			connections.Add(new Connection($"TYO SW[{color}]-L02", "Ethernet 1", $"TYO SW[{color}]-S01", "Ethernet 50"));
			connections.Add(new Connection($"TYO SW[{color}]-L02", "Ethernet 2", $"TYO SW[{color}]-S02", "Ethernet 50"));
			connections.Add(new Connection($"TYO SW[{color}]-L03", "Ethernet 1", $"TYO SW[{color}]-S01", "Ethernet 48"));
			connections.Add(new Connection($"TYO SW[{color}]-L03", "Ethernet 2", $"TYO SW[{color}]-S02", "Ethernet 48"));
		}

		connections.Add(new Connection("NY SW-L01", "Ethernet 1", "NY SW-S01", "Ethernet 1"));
		connections.Add(new Connection("NY SW-L02", "Ethernet 1", "NY SW-S01", "Ethernet 2"));
		connections.Add(new Connection("NY SW-L03", "Ethernet 1", "NY SW-S01", "Ethernet 3"));
		connections.Add(new Connection("NY SW-L04", "Ethernet 1", "NY SW-S01", "Ethernet 4"));
		connections.Add(new Connection("NY SW-L01", "Ethernet 2", "NY SW-S02", "Ethernet 1"));
		connections.Add(new Connection("NY SW-L02", "Ethernet 2", "NY SW-S02", "Ethernet 2"));
		connections.Add(new Connection("NY SW-L03", "Ethernet 2", "NY SW-S02", "Ethernet 3"));
		connections.Add(new Connection("NY SW-L04", "Ethernet 2", "NY SW-S02", "Ethernet 4"));

		connections.Add(new Connection("LA SW[BLUE]-S01", "Ethernet 1", "LA SW[BLUE]-S02", "Ethernet 1"));
		connections.Add(new Connection("TYO SW[BLUE]-S01", "Ethernet 1", "TYO SW[BLUE]-S02", "Ethernet 1"));
		connections.Add(new Connection("LA SW[RED]-S01", "Ethernet 1", "LA SW[RED]-S02", "Ethernet 1"));
		connections.Add(new Connection("TYO SW[RED]-S01", "Ethernet 1", "TYO SW[RED]-S02", "Ethernet 1"));

		connections.Add(new Connection("LA SW[BLUE]-S01", "Ethernet 2", "TYO SW[BLUE]-S01", "Ethernet 2"));
		connections.Add(new Connection("LA SW[RED]-S01", "Ethernet 2", "TYO SW[RED]-S01", "Ethernet 2"));
		connections.Add(new Connection("LA SW[BLUE]-S02", "Ethernet 2", "NY SW-S01", "Ethernet 6"));
		connections.Add(new Connection("LA SW[RED]-S02", "Ethernet 2", "NY SW-S02", "Ethernet 6"));
		connections.Add(new Connection("TYO SW[BLUE]-S02", "Ethernet 2", "NY SW-S02", "Ethernet 7"));
		connections.Add(new Connection("TYO SW[RED]-S02", "Ethernet 2", "NY SW-S01", "Ethernet 7"));
		connections.Add(new Connection("NY SW-S01", "Ethernet 8", "NY SW-S02", "Ethernet 8"));

		connections.Add(new Connection("NY SW-S01", "Ethernet 9", "SG SW[BLUE]-S01", "Ethernet 1"));
		connections.Add(new Connection("NY SW-S02", "Ethernet 9", "SG SW[RED]-S01", "Ethernet 1"));
		connections.Add(new Connection("PAR SW[BLUE]-S01", "Ethernet 1", "SG SW[BLUE]-S01", "Ethernet 2"));
		connections.Add(new Connection("PAR SW[RED]-S01", "Ethernet 1", "SG SW[RED]-S01", "Ethernet 2"));
		connections.Add(new Connection("SG SW[BLUE]-S01", "Ethernet 3", "SG SW[RED]-S01", "Ethernet 3"));

		connections.Add(new Connection("LA SW[BLUE]-S01", "Ethernet 3", "OB SW[BLUE]-S01", "Ethernet 1"));
		connections.Add(new Connection("LA SW[RED]-S01", "Ethernet 3", "OB SW[RED]-S01", "Ethernet 1"));
		connections.Add(new Connection("OB SW[BLUE]-S01", "Ethernet 2", "OB SW[RED]-S01", "Ethernet 2"));
		connections.Add(new Connection("OB SW[BLUE]-S01", "Ethernet 3", "OB SW-L01", "Ethernet 1"));
		connections.Add(new Connection("OB SW[BLUE]-S01", "Ethernet 4", "OB SW-L02", "Ethernet 1"));
		connections.Add(new Connection("OB SW[RED]-S01", "Ethernet 3", "OB SW-L01", "Ethernet 2"));
		connections.Add(new Connection("OB SW[RED]-S01", "Ethernet 4", "OB SW-L02", "Ethernet 2"));

		// Inter-region BC links: NYC core BCs to LA access BCs
		connections.Add(new Connection("NY SW-L01", "Ethernet 10", "LA SW[BLUE]-L01", "Ethernet 10"));
		connections.Add(new Connection("NY SW-L01", "Ethernet 11", "LA SW[BLUE]-L02", "Ethernet 10"));
		connections.Add(new Connection("NY SW-L01", "Ethernet 12", "LA SW[BLUE]-L03", "Ethernet 10"));
		connections.Add(new Connection("NY SW-L02", "Ethernet 10", "LA SW[RED]-L01", "Ethernet 10"));
		connections.Add(new Connection("NY SW-L02", "Ethernet 11", "LA SW[RED]-L02", "Ethernet 10"));
		connections.Add(new Connection("NY SW-L02", "Ethernet 12", "LA SW[RED]-L03", "Ethernet 10"));

		// LA access BCs to OB edge BCs
		connections.Add(new Connection("LA SW[BLUE]-L01", "Ethernet 11", "OB SW-L01", "Ethernet 10"));
		connections.Add(new Connection("LA SW[BLUE]-L02", "Ethernet 11", "OB SW-L02", "Ethernet 10"));

		return connections;
	}

	private static int CreateViewIfNotExists(IDms dms, string viewName, int parentViewID = -1)
	{
		if (!dms.ViewExists(viewName))
		{
			int newViewID = dms.CreateView(new ViewConfiguration(viewName, dms.GetView(parentViewID)));
			int retries = 0;
			while (!dms.ViewExists(newViewID) && retries < 20)
			{
				Thread.Sleep(100);
				retries++;
			}
			return newViewID;
		}
		return dms.GetView(viewName).Id;
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
					engine.GenerateInformation($"Protocol '{entry.Name}' version '{version}' set as Production.");
				}
				catch (Exception ex)
				{
					engine.GenerateInformation($"Warning: Failed to set '{entry.Name}': {ex.Message}");
				}
			}
		}
		catch (Exception ex)
		{
			engine.GenerateInformation($"Warning: Exception setting protocol versions: {ex.Message}");
		}
	}

	private static string GenerateClockId(int numericId)
	{
		IEnumerable<string> idBytes = BitConverter.GetBytes(numericId).Reverse().Select(b => b.ToString("X2"));
		const int IntermediaOctets = 1;
		return $"01:BA:{String.Join(":", idBytes.Take(IntermediaOctets))}:FF:FE:{String.Join(":", idBytes.Skip(IntermediaOctets))}";
	}

	private class ElementConfig
	{
		public ElementConfig(string name, Dictionary<int, object> parametersToSet)
		{
			ElementName = name;
			ParametersToSet = new Dictionary<int, object>(parametersToSet);
		}

		public string ElementName { get; set; }
		public Dictionary<int, object> ParametersToSet { get; set; }
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

		public Connection(string sourceElement, string sourceInterfaceName, string destinationElement, string destinationInterfaceName)
		{
			SourceElementName = sourceElement;
			SourceInterfaceName = sourceInterfaceName;
			DestinationElementName = destinationElement;
			DestinationInterfaceName = destinationInterfaceName;
		}

		public void MakeConnection(IEngine engine)
		{
			try
			{
				SourceElement = engine.FindElement(SourceElementName);
				if (SourceElement == null)
				{
					engine.GenerateInformation($"Warning: Source element '{SourceElementName}' not found");
					return;
				}

				DestinationElement = engine.FindElement(DestinationElementName);
				if (DestinationElement == null)
				{
					engine.GenerateInformation($"Warning: Destination element '{DestinationElementName}' not found");
					return;
				}

				SourceInterfaceID = GetInterfaceID(engine, SourceElement, SourceInterfaceName);
				DestinationInterfaceID = GetInterfaceID(engine, DestinationElement, DestinationInterfaceName);

				if (SourceInterfaceID == 0 || DestinationInterfaceID == 0)
				{
					engine.GenerateInformation($"Warning: Interface lookup failed");
					return;
				}

				GetConnectionNames();

				if (!Exists(engine))
				{
					var msg = new EditConnection()
					{
						Action = ConnectivityEditAction.Add,
						ConnectionID = 0,
						DataMinerID = SourceElement.DmaId,
						HostingDataMinerID = SourceElement.DmaId,
						EditBothConnections = true,
						SourceElementID = SourceElement.ElementId,
						SourceInterfaceID = SourceInterfaceID,
						SourceName = GetSourceConnectionName(),
						DestinationElement = DestinationElement.DmaId + "/" + DestinationElement.ElementId,
						DestinationInterfaceID = DestinationInterfaceID,
						DestinationName = GetDestinationConnectionName(),
					};
					engine.SendSLNetMessage(msg);
				}
			}
			catch (Exception e)
			{
				engine.GenerateInformation($"Error creating connection: {e.Message}");
			}
		}

		private bool Exists(IEngine engine)
		{
			try
			{
				string[] keys = SourceElement.GetTablePrimaryKeys(65060);
				if (keys == null) return false;

				foreach (string key in keys)
				{
					int srcID = Convert.ToInt32(SourceElement.GetParameter(65062, key));
					if (srcID == SourceInterfaceID)
					{
						string destElem = Convert.ToString(SourceElement.GetParameter(65089, key));
						if (destElem == DestinationElement.DmaId + "/" + DestinationElement.ElementId)
						{
							int destID = Convert.ToInt32(SourceElement.GetParameter(65064, key));
							if (destID == DestinationInterfaceID)
								return true;
						}
					}
				}
			}
			catch { }
			return false;
		}

		private int GetInterfaceID(IEngine engine, Element element, string name)
		{
			try
			{
				Interface[] interfaces = element.GetInterfaces();
				if (interfaces == null) return 0;

				foreach (Interface iface in interfaces)
				{
					if (iface.Name == name)
						return iface.InterfaceId;
				}

				foreach (Interface iface in interfaces)
				{
					if (iface.Name.Contains(name))
						return iface.InterfaceId;
				}
			}
			catch { }
			return 0;
		}

		private string GetSourceConnectionName() => sourceConnectionNameBuffer ?? "Unknown";
		private string GetDestinationConnectionName() => destinationConnectionNameBuffer ?? "Unknown";

		private void GetConnectionNames()
		{
			try
			{
				string src = Convert.ToString(SourceElement.GetParameter(65093, Convert.ToString(SourceInterfaceID)));
				string dest = Convert.ToString(DestinationElement.GetParameter(65093, Convert.ToString(DestinationInterfaceID)));
				sourceConnectionNameBuffer = $"{src} -> {dest}";
				destinationConnectionNameBuffer = $"{dest} -> {src}";
			}
			catch
			{
				sourceConnectionNameBuffer = $"Ethernet {SourceInterfaceID}";
				destinationConnectionNameBuffer = $"Ethernet {DestinationInterfaceID}";
			}
		}
	}
}







