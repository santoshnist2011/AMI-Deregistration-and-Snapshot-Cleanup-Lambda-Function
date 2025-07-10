# AMI-Deregistration-and-Snapshot-Cleanup-Lambda-Function
This AWS Lambda function identifies and deregisters Amazon Machine Images (AMIs) tagged with Purpose: Patching and Delete: Yes that are older than a week, along with their associated snapshots. It sends a report via Amazon SES to specified email recipients
