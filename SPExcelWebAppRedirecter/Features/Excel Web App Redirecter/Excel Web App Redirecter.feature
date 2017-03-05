﻿<?xml version="1.0" encoding="utf-8"?>
<feature xmlns:dm0="http://schemas.microsoft.com/VisualStudio/2008/DslTools/Core" dslVersion="1.0.0.0" Id="851e5539-f0c7-462a-a3c9-5f46b93e8d15" description="This feature gives you extra flexibility in terms of choosing Excel Services or Excel Web App for viewing in browser mode.&#xD;&#xA;In case you are using Office Web Apps and your SharePoint farm is configured to support Business Intelligence functionality with Excel, you most likely configured it using: New-SPWOPISuppressionSetting -Extension XLSX -Action VIEW; and New-SPWOPISuppressionSetting -Extension XLS -Action VIEW;. This is making Excel Services to become the default view solution for Excel files. However, this is a farm-wide setting (you don't have too much flexibility there) and that means Excel Web App will never be used for view mode. Activate this feature in order to change this behavior and redirect any Excel Services request to Excel Web App (part of Office Web Apps)." featureId="851e5539-f0c7-462a-a3c9-5f46b93e8d15" imageUrl="" receiverAssembly="$SharePoint.Project.AssemblyFullName$" receiverClass="$SharePoint.Type.23d91ccd-7e4a-4464-9f20-efc3ab15d274.FullName$" scope="WebApplication" solutionId="00000000-0000-0000-0000-000000000000" title="Excel Web App Redirecter" version="" deploymentPath="$SharePoint.Project.FileNameWithoutExtension$_$SharePoint.Feature.FileNameWithoutExtension$" xmlns="http://schemas.microsoft.com/VisualStudio/2008/SharePointTools/FeatureModel" />