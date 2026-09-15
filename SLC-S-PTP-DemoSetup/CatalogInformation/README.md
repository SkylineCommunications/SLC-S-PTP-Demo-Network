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
- **Required connectors**:
  - [Generic Switch](https://catalog.dataminer.services/details/ae750d4e-50b1-489d-b948-1abe1af591dd)
  - [Generic Edge Chassis](https://catalog.dataminer.services/details/c5b2be4c-ecc2-4b59-a274-038bed39274f)
- **Recommended connector**:
  - [Skyline PTP](https://catalog.dataminer.services/details/fb19f440-e363-429a-8aed-e756dd4c3c48)
- **Recommended standard solution**:
  - [PTP](https://catalog.dataminer.services/details/9c5eb0a1-43bc-42d2-bca2-de4982ee57d7)

## Technical Reference

- [PTP Standard Solution Documentation](https://docs.dataminer.services/solutions/standard_solutions/PTP/SolPTP.html)
- [Contacting DataMiner Support](https://aka.dataminer.services/contacting-tech-support)