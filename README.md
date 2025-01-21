Craigslist Automation Process Flow Using Selenium Grid

Prerequisites
1. Install the following tools:
   - Java JDK
   - Python

2. Setup Machines:
   - Machine A (Hub)
   - Machine B (Node)
   (Supports n number of nodes.)

3. Configure ZeroTier Network:
   - Connect both machines using a ZeroTier network.
   - Assign Managed IPs through ZeroTier:
     - Example IP for Machine A: 192.168.196.159
     - Example IP for Machine B: 192.168.196.44


Step 1: Configure Machine A (Hub)

1. Edit config.TOML File:
   - Set the host parameter to the IP of Machine A: 192.168.196.159.
   - Update the description field to 400 (visible on ZeroTier).

2. Edit startHub.bat File:
   - Add the IP address of Machine A (192.168.196.159).

3. Start the Selenium Hub:
   - Run startHub.bat.
   - Note the returned URL:
     http://192.168.196.159:4444
     (This URL confirms the grid is established.)


Step 2: Configure Machine B (Node)

1. Edit config.TOML File:
   - Set the host parameter to the IP of Machine B: 192.168.196.44.
   - Update the description field to 203 (visible on ZeroTier).

2. Edit startNode.bat File:
   - Add the Hub's IP and port:
     --hub http://192.168.196.159:4444.

3. Start the Selenium Node:
   - Run startNode.bat.

4. Verify Node Addition:
   - Open the Selenium Grid URL:
     http://192.168.196.159:4444
     (The node should be listed on the grid.)


Step 3: Add Additional Nodes (Optional)

1. Repeat the Machine B Configuration for other machines.
2. Ensure each node has a unique IP assigned through ZeroTier.
3. All nodes should point to the Hub's IP in their startNode.bat file.


Step 4: Run the Python Script

1. Prepare the Script:
   - Open the Python script in PyCharm.
   - Install all required dependencies and libraries using pip.

2. Configure Python Environment:
   - Add the Python interpreter if not already configured.

3. Run the Script:
   - Execute the script in PyCharm.

4. Troubleshooting:
   - If the script does not run:
     - Verify the installed version of Google Chrome.
     - In the Selenium Grid folder, update the WebDriver to match the Chrome version.


This process establishes a Selenium Grid with Machine A as the Hub and Machine B (and optionally additional machines) as Nodes. The Python script utilizes this grid for Craigslist automation.
