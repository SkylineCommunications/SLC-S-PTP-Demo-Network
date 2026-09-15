# PTP Demo Network Setup

## About

The **PTP Demo Network Setup** package provisions a realistic, multi-site Precision Time Protocol (PTP) simulation environment within DataMiner. It is designed to demonstrate and validate the **PTP Standard Solution** (`SLC-S-PTP`) without requiring access to live physical broadcast or telecom timing hardware.

The solution automatically sets up a complete multi-domain PTP clock topology spanning multiple geographic locations (New York, Tokyo, Los Angeles, Singapore, Paris, and an Outside Broadcast truck), complete with physical DataMiner Connectivity Framework (DCF) interface links and IEEE 1588 hierarchical clock configurations.

## Key Features

- **Simulate Multi-Site PTP Topology**: Deploys simulated Grandmasters, Boundary Clocks, Transparent Clocks, and Follower devices across multiple interconnected regions.
- **Model Physical DCF Cabling**: Automatically configures full DCF interface links representing physical network cabling between all switches.
- **Establish Hierarchical PTP Sync**: Generates valid IEEE 1588 parent clock relationships, clock identities, and steps-removed metadata along DCF paths.
- **Provision Multi-Domain Clocks**: Configures devices across both primary and test PTP domains (Domain 101 and 111) for multi-domain monitoring scenarios.
- **Accelerate Demonstration & Testing**: Eliminates manual device setup and SNMP simulation wiring so operators and engineers can test PTP monitoring immediately.

## Use Cases

- **PTP Solution Demonstrations**: Showcase the capabilities of the DataMiner PTP Standard Solution to stakeholders and customers without live broadcast infrastructure.
- **Topology & Visualization Testing**: Validate PTP topology visualization, auto-arrange behaviors, and DCF versus Parent Clock connection modes on a rich network graph.
- **Operator Training**: Familiarize operations teams with PTP clock class variations, jitter monitoring, and clock hierarchy inspection in a safe demo environment.

## Prerequisites

- **DataMiner**: Version 10.3.0.0 or higher
- **Dependencies**: None. The `SLC-S-PTP` standard solution is recommended to visualize and monitor the demo environment, but it can be installed either before or after this demo package.

## Technical Reference

- [PTP Standard Solution Documentation](https://docs.dataminer.services/solutions/standard_solutions/PTP/SolPTP.html)
- [Contacting DataMiner Support](https://aka.dataminer.services/contacting-tech-support)