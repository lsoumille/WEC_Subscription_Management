# WEC_Subscription_Management
Powershell script used for WEC scalability

## WEC Limitations

There are three factors that limit the scalability of WEC servers. The general rule for a stable WEC server on commodity hardware is “10k x 10k” – meaning, no more than 10,000 concurrently active WEF Clients per WEC server and no more than 10,000 events/second average event volume.

Disk I/O. The WEC server does not process or validate the received event, but rather buffers the received event and then logs it to a local event log file (EVTX file). The speed of logging to the EVTX file is limited by the disk write speed. Isolating the EVTX file to its own array or using high speed disks can increase the number of events per second that a single WEC server can receive.

Network Connections. While a WEF source does not maintain a permanent, persistent connection to the WEC server, it does not immediately disconnect after sending its events. This means that the number of WEF sources that can simultaneously connect to the WEC server is limited to the open TCP ports available on the WEC server.

Registry size. For each unique device that connects to a WEF subscription, there is a registry key (corresponding to the FQDN of the WEF Client) created to store bookmark and source heartbeat information. If this is not pruned to remove inactive clients this set of registry keys can grow to an unmanageable size over time.

When a subscription has >1000 WEF sources connect to it over its operational lifetime, also known as lifetime WEF sources, Event Viewer can become unresponsive for a few minutes when selecting the Subscriptions node in the left-navigation, but will function normally afterwards.
At >50,000 lifetime WEF sources, Event Viewer is no longer an option and wecutil.exe (included with Windows) must be used to configure and manage subscriptions.
At >100,000 lifetime WEF sources, the registry will not be readable and the WEC server will likely have to be rebuilt.

## Solution

Use this script to merge all WEC subscriptions into in big one to reduce WEF sources
