# Erlang C Calculator VBA Module

## Overview
The Erlang C Calculator is a module designed to assist with call center operations, allowing users to calculate various performance metrics based on arrival rates, service rates, and the number of agents.

## Features
- **Arrival Rate Calculation**: Determine the rate at which calls are received.
- **Service Rate Calculation**: Estimate the rate at which calls can be handled by agents.
- **Agent Configuration**: Input the number of agents available for service.
- **Performance Metrics Summation**: Calculate key statistics including average wait time, probability of wait, and service level.
- **User-Friendly Interface**: Simple input forms for ease of use.

## Installation
1. **Download the Module**: Clone or download the repository from GitHub.
2. **Import the VBA Module**: Open Excel, press `ALT` + `F11` to open the VBA editor, and import the module using `File -> Import File`.
3. **Enable Macros**: Ensure that macros are enabled in your Excel settings.

## Usage Examples
- **Basic Calculation**: 
  ```vba
  Dim result As Double
  result = CalculateErlangC(arrivalRate, serviceRate, numberOfAgents)
  MsgBox "Expected wait time: " & result
  ```
- **Agent Configuration**:
  Setup the number of agents dynamically based on user input using a form.

## Function Reference
- **CalculateErlangC(arrivalRate As Double, serviceRate As Double, numberOfAgents As Integer) As Double**
  - **Parameters**:
    - `arrivalRate`: The rate of incoming calls.
    - `serviceRate`: The rate of service provided by each agent.
    - `numberOfAgents`: The total number of agents available.
  - **Returns**: A double representing the expected wait time in seconds.

## Assumptions
- Calls arrive following a Poisson distribution.
- Agents serve customers at a constant average rate.
- The calculator is suitable for steady-state operations.

## Technical Notes
- The module is designed to handle a range of operating scenarios. 
- Optimization for larger call volumes may require fine-tuning of service rates.

## Author
This module was developed by Mike-Launizar. Contributions and feedback are welcome! 

## License
This project is licensed under the MIT License. See the LICENSE file for details.