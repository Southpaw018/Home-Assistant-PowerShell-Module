# Create default configuration
if(-Not (Test-Path -Path $PSScriptRoot\Config.xml)) 
{
  # Create the configuration object
  $ConfigObject = @{
    Service  = 'input_boolean.toggle'
    Entity   = 'input_boolean.idle'
    HostName = 'hass.mydomain.com'
    Port     = '8123'
    IdleSeconds = '10'
    CheckInterval = '5'
    Token    = ''
  }

  # Export the object to XML
  $ConfigObject | Export-Clixml -Path "$PSScriptRoot\Config.xml" -Force
  break
}

# Test if config file exists and import it
if(Test-Path -Path "$PSScriptRoot\Config.xml") {
  
  # Import Configuration
  $config = Import-Clixml -Path "$PSScriptRoot\Config.xml"
} else {

  Write-Error -Message "Configuration file $PSScriptRoot\Config.xml not found"
  break
}


Function Get-AudioPlaying 
{
  Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace Foo
{
    public class Bar
    {
        public static bool IsWindowsPlayingSound()
        {
            IMMDeviceEnumerator enumerator = (IMMDeviceEnumerator)(new MMDeviceEnumerator());
            IMMDevice speakers = enumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eMultimedia);
            IAudioMeterInformation meter = (IAudioMeterInformation)speakers.Activate(typeof(IAudioMeterInformation).GUID, 0, IntPtr.Zero);
            float value = meter.GetPeakValue();

            // this is a bit tricky. 0 is the official "no sound" value
            // but for example, if you open a video and plays/stops with it (w/o killing the app/window/stream),
            // the value will not be zero, but something really small (around 1E-09)
            // so, depending on your context, it is up to you to decide
            // if you want to test for 0 or for a small value
            return value > 0;
        }

        [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
        private class MMDeviceEnumerator
        {
        }

        private enum EDataFlow
        {
            eRender,
            eCapture,
            eAll,
        }

        private enum ERole
        {
            eConsole,
            eMultimedia,
            eCommunications,
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("A95664D2-9614-4F35-A746-DE8DB63617E6")]
        private interface IMMDeviceEnumerator
        {
            void NotNeeded();
            IMMDevice GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role);
            // the rest is not defined/needed
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("D666063F-1587-4E43-81F1-B948E807363F")]
        private interface IMMDevice
        {
            [return: MarshalAs(UnmanagedType.IUnknown)]
            object Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams);
            // the rest is not defined/needed
        }

        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("C02216F6-8C67-4B5B-9D00-D008E73E0064")]
        private interface IAudioMeterInformation
        {
            float GetPeakValue();
            // the rest is not defined/needed
        }
    }
}
'@

  $obj = [Foo.Bar]::IsWindowsPlayingSound()

  Write-Output -InputObject $obj
  Remove-Variable -Name obj
}


Function Get-IdleTime 
{
  Add-Type -TypeDefinition @'
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace PInvoke.Win32 {

    public static class UserInput {

        [DllImport("user32.dll", SetLastError=false)]
        private static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

        [StructLayout(LayoutKind.Sequential)]
        private struct LASTINPUTINFO {
            public uint cbSize;
            public int dwTime;
        }

        public static DateTime LastInput {
            get {
                DateTime bootTime = DateTime.UtcNow.AddMilliseconds(-Environment.TickCount);
                DateTime lastInput = bootTime.AddMilliseconds(LastInputTicks);
                return lastInput;
            }
        }

        public static TimeSpan IdleTime {
            get {
                return DateTime.UtcNow.Subtract(LastInput);
            }
        }

        public static int LastInputTicks {
            get {
                LASTINPUTINFO lii = new LASTINPUTINFO();
                lii.cbSize = (uint)Marshal.SizeOf(typeof(LASTINPUTINFO));
                GetLastInputInfo(ref lii);
                return lii.dwTime;
            }
        }
    }
}
'@

  $Idle = [PInvoke.Win32.UserInput]::IdleTime

  $obj = [PSCustomObject]@{
    Days    = $Idle.Days
    Hours   = $Idle.Hours
    Minutes = $Idle.Minutes
    Seconds = $Idle.Seconds
  }

  Write-Output -InputObject $obj
  Remove-Variable -Name obj
}

Function Update-HomeAssistantEntity
{
  [CmdletBinding()]
  param(
    [ValidateSet('on','off')]
    [string[]]
    $State
  )

  # Import the home assistant module
  Import-Module -Name $PSScriptRoot\Home-Assistant.psd1 -Force

  # Check if an existing session exists else create it
  if(-Not ($ha_api_configured)) 
  {
    New-HomeAssistantSession -Hostname $config.HostName -Port $config.Port -Token $config.Token -UseSSL
  }

  # Get the entity current state
  $CurrentState = (Get-HomeAssistantEntity -entity_id $config.Entity).state

  #If we are set setting the entitiy to 'off' check that it is actually not equal to off first
  if($State -eq 'off') 
  {
    if($CurrentState -ne 'off') 
    {
      $null = Invoke-HomeAssistantService -service $config.Service -entity_id $config.Entity
    }
  }

  #If we are set setting the entitiy to 'on' check that it is actually not equal to on first
  if($State -eq 'on') 
  {
    if($CurrentState -ne 'on') 
    {
      $null = Invoke-HomeAssistantService -service $config.Service -entity_id $config.Entity
    }
  }

  # Output the current status
  $obj = (Get-HomeAssistantEntity -entity_id $config.Entity).state

  Write-Output -InputObject "Service: $($config.Service) - Entity: $($config.Entity) - State: $($obj)"
}

# Loop the audio playing and idle status and then update the entity accordingly
while($true)
{
  if((Get-IdleTime).Seconds -gt $config.IdleSeconds -and (Get-AudioPlaying) -eq $false) 
  {
    Update-HomeAssistantEntity -State on
  }
  else 
  {
    Update-HomeAssistantEntity -State off
  }

  Start-Sleep -Seconds $config.CheckInterval
}

