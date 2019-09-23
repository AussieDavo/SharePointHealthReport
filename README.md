# SharePointHealthReport
SharePoint Server Health Report

This PowerShell script runs from a single SharePoint server, it collects information from the entire SharePoint Farm and sends the report via email.

Script run time: 30 seconds (on average)

Compatibility: SharePoint 2016 and 2013 (not yet tested on earlier versions on SharePoint or SharePoint 2019)

- IIS Application Pool Status
- IIS Web Site Status
  - Includes certificate bindings with expiry dates
- Service Application and Proxy Status
- SharePoint Health Analyser Reports
- Distributed Cache Health
  - Includes Cache Size
- Search Topology Health
- Services on Server (Service Instances) Status
  - Includes the ability to capture a configuration baseline and monitor configuration drift
- Crawl Log - Crawl History (Last 4 hours)
  - Includes Average Crawl Duration over the last 24 hours
  - Includes Average Crawl Rate over the last 24 hours
  - Includes total searchable items
- Crawl Log - Error Breakdown
  - Includes Total Crawl Errors
