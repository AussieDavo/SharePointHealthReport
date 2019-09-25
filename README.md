# SharePointHealthReport
SharePoint Server Health Report

This PowerShell script runs from a single SharePoint server, it collects information from the entire SharePoint Farm and sends the report via email.

Script run time: 30 seconds (on average)

Compatibility: SharePoint 2016 and 2013 (not yet tested on earlier versions of SharePoint or SharePoint 2019)

- IIS Application Pool Status
- IIS Web Site Status
  - Includes certificate bindings with expiry dates
- Service Application and Proxy Status
- SharePoint Health Analyser Reports
- Distributed Cache Health
  - Includes Cache Size
  - Copies AppFabric configuration back from remote Distributed Cache Server's registry in order to get cache cluster health locally 
- Search Topology Health
- Services on Server (Service Instances) Status
  - Includes the ability to capture a configuration baseline and monitor configuration drift
- Crawl Log - Crawl History (Last 4 hours)
  - Includes Average Crawl Duration over the last 24 hours
  - Includes Average Crawl Rate over the last 24 hours
  - Includes total searchable items
- Crawl Log - Error Breakdown
  - Includes Total Crawl Errors


![IIS App Pool Status](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/IIS%20App%20Pool%20Status.jpg)

![IIS Web Site Status](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/IIS%20Web%20Site%20Status.jpg)

![Service Application and Proxy Status](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/Service%20Application%20and%20Proxy%20Status.jpg)

![SharePoint Health Analyser Reports](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/SharePoint%20Health%20Analyser%20Reports.jpg)

![Distributed Cache Health](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/Distributed%20Cache%20Health.jpg)

![Search Topology Health](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/Search%20Topology%20Health.jpg)

![Services on Server - Service Instances](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/Services%20on%20Server%20-%20Service%20Instances.jpg)

![Crawl Log - Crawl History - Last 24 hours](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/Crawl%20Log%20-%20Crawl%20History%20-%20Last%2024%20hours.jpg)

![Crawl Log - Error Breakdown](https://raw.githubusercontent.com/AussieDavo/SharePointHealthReport/master/_images/Crawl%20Log%20-%20Error%20Breakdown.jpg)


