### Author: Iftode Viorel
### Description: This PowerShell scripts helps you to deploy SharePoint WSP solutions.
### Usage: Save this script in the same folder with SPExcelWebAppRedirecter.wsp. Execute 
### the script using .\deploySPExcelWebAppRedirecter.ps1 and this will deploy the 
### solution into your SharePoint farm.

if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -eq $null) {
    Add-PSSnapin "Microsoft.SharePoint.PowerShell" > $null;
}

$WSPDeploymentFolderPath=(([System.IO.Directory]::GetParent($MyInvocation.MyCommand.Path)).FullName);

##function used to retract SharePoint custom solutions
##
##Example:
##uninstallSPsolution -WSPName $WSPName -AllWebApplications:$AllWebApplications;
function uninstallSPSolution{
	param(
		[Parameter(Mandatory=$True,Position=1)]
		[string]$WSPName,
		[boolean]$AllWebApplications=$false
	)
	$SPSolution = Get-SPSolution -identity $WSPName -ErrorAction:SilentlyContinue;
	if($SPSolution -ne $null)
	{
		Write-Host 'Retracting' $WSPName -ForegroundColor:Green -BackgroundColor:Black;
        $command='Uninstall-SPSolution -identity "'+$WSPName+'" -confirm:$false';
		if($AllWebApplications -eq $True)
		{
			$command+=' -AllWebApplications';
		}
		Write-Host 'Executing >'$command -ForegroundColor:Green -BackgroundColor:Black;
        Invoke-Expression $command;
		Write-Host 'Retracting...' -ForegroundColor:Green -BackgroundColor:Black;
		do {Write-Host . -ForegroundColor:Green -BackgroundColor:Black -NoNewline; Start-Sleep 2;} while ($SPSolution.JobExists);
		Write-Host;
	}
	else
	{
		Write-Host 'Couldn''t found the solution' $WSPName 'deployed on the farm!' -ForegroundColor:Red -BackgroundColor:Black;
	}
}


##function used to remove* SharePoint custom solutions
## *it first retracts the solution
##
##Example:
##removeSPSolution -WSPName 'SolutionName.wsp' -AllWebApplications:$true;
function removeSPSolution{
	param(
		[Parameter(Mandatory=$True,Position=1)]
		[string]$WSPName,
		[boolean]$AllWebApplications=$false,
		[boolean]$Force=$false
	)
    Write-Host '-----------------------------' -ForegroundColor:Yellow -BackgroundColor:Black;
	$SPSolution = Get-SPSolution -identity $WSPName -ErrorAction:SilentlyContinue;
	if($SPSolution -ne $null)
	{
		if($SPSolution.Deployed -eq $True)
		{
			uninstallSPsolution -WSPName $WSPName -AllWebApplications:$AllWebApplications;
		}
		Write-Host 'Removing' $WSPName -ForegroundColor:Green -BackgroundColor:Black;
        $command='Remove-SPSolution -identity "'+$WSPName+'" -confirm:$false';
        if($Force)
        {
            $command+=' -Force';
        }
        Write-Host 'Executing >'$command -ForegroundColor:Green -BackgroundColor:Black;
        Invoke-Expression $command;
		Write-Host 'Removed!' -ForegroundColor:Green -BackgroundColor:Black;
	}
	else
	{
		Write-Host 'Couldn''t found the solution' $WSPName 'deployed on the farm!' -ForegroundColor:Red -BackgroundColor:Black;
	}
    Write-Host '-----------------------------' -ForegroundColor:Yellow -BackgroundColor:Black;
}


##function used to deploy SharePoint WSP solutions
##
##Example:
##deploySPSolution -WSPName 'SolutionName.wsp' -GACDeployment:$True -AllWebApplications:$True -WSPDeploymentFolder 'C:\FolderWhereTheWSPfileIsHosted';
##
##or copy this script in the same folder with the WSP files and the call will be
##deploySPSolution -WSPName 'SolutionName.wsp' -GACDeployment:$True -AllWebApplications:$True;
function deploySPSolution{
    param(
		[Parameter(Mandatory=$True,Position=1)]
		[string]$WSPName,
        [boolean]$GACDeployment=$false,
		[boolean]$AllWebApplications=$false,
		[boolean]$Force=$false,
        [string]$WSPDeploymentFolder=[String]::Empty
	)
    Write-Host '-----------------------------' -ForegroundColor:Yellow -BackgroundColor:Black;
    $WSPFilePath=[String]::Empty;
    if($WSPDeploymentFolder -eq [String]::Empty)
    {
        $WSPFilePath=$WSPDeploymentFolderPath+'\'+$WSPName;
    }
    else
    {
        if([System.IO.Directory]::Exists($WSPDeploymentFolder))
        {
			$Folder=New-Object System.IO.DirectoryInfo($WSPDeploymentFolder);
			if($Folder.FullName -eq $Folder.Root)
			{
				$WSPFilePath=$Folder.FullName+$WSPName;
			}
			else
			{
				$WSPFilePath=$Folder.Parent.FullName+'\'+$Folder.Name+'\'+$WSPName;
			}
        }
        else
        {
            Write-Host 'Invalid WSPDeploymentFolder path!' -ForegroundColor:Red -BackgroundColor:Black;
        }
    }
    
    if([System.IO.File]::Exists($WSPFilePath))
    {
        Write-Host $WSPName 'has been found in the deployment folder.' -ForegroundColor:Green -BackgroundColor:Black;
        
        $command='Add-SPSolution -LiteralPath "'+$WSPFilePath+'" -confirm:$false';
        Write-Host 'Executing >'$command -ForegroundColor:Green -BackgroundColor:Black;
        Invoke-Expression $command;
        
        $command='Install-SPSolution -Identity "'+$WSPName+'" -confirm:$false';
        if($GACDeployment)
        {
            $command+=' -GACDeployment';
        }
        if($AllWebApplications)
        {
            $command+=' -AllWebApplications';
        }
        if($Force)
        {
            $command+=' -Force';
        }
        
        Write-Host 'Executing >'$command -ForegroundColor:Green -BackgroundColor:Black;
        Invoke-Expression $command;
        $SPSolution = Get-SPSolution -identity $WSPName -ErrorAction:SilentlyContinue;
        if($SPSolution -ne $null)
	    {
            Write-Host 'Deploying...' -ForegroundColor:Green -BackgroundColor:Black;
	    	do {Write-Host . -ForegroundColor:Green -BackgroundColor:Black -NoNewline; Start-Sleep 2;} while ($SPSolution.JobExists);
	    	Write-Host;
        }
    }
    else
    {
        Write-Host 'I couldn''t found' $WSPFilePath'!' -ForegroundColor:Red -BackgroundColor:Black;
    }
    Write-Host '-----------------------------' -ForegroundColor:Yellow -BackgroundColor:Black;
}



##------------------------remove------------------------##
removeSPSolution -WSPName 'SPExcelWebAppRedirecter.wsp' -Force:$True;
##------------------------remove------------------------##

##------------------------deploy------------------------##
deploySPSolution -WSPName 'SPExcelWebAppRedirecter.wsp' -GACDeployment:$True -Force:$True;
##------------------------deploy------------------------##

    
trap
{
	Write-Host "===========================" -ForegroundColor Green;
	Write-Host "Something weird happened. Please contact the SharePoint admin and tell him step by step what did you do to receive this message." -ForegroundColor:Red -BackgroundColor:Black;
	Write-Host "===========================" -ForegroundColor Green;
}

if ((Get-PSSnapin "Microsoft.SharePoint.PowerShell" -ErrorAction SilentlyContinue) -ne $null) {
    Remove-PSSnapin "Microsoft.SharePoint.PowerShell" > $null;
}
