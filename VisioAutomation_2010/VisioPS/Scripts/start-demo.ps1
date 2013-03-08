# Version: 1.1
function Start-Demo
{
  param($file=".\demo.txt", [int]$command=0)
  $CommentColor = "Yellow"
  $MetaCommandColor = "Red"
  Clear-Host

  $_Random = New-Object System.Random
  $_lines = @(Get-Content $file)
  $_starttime = [DateTime]::now
  $_PretendTyping = $true    
  $_InterkeyPause = 200
  Write-Host -for $CommentColor @"
NOTE: Start-Demo replaces the typing but runs the actual commands.
.
<Demo [$file] Started.  Type `"?`" for help>
"@


  # We use a FOR and an INDEX ($_i) instead of a FOREACH because
  # it is possible to start at a different location and/or jump 
  # around in the order.
  for ($_i = $Command; $_i -lt $_lines.count; $_i++)
  {  
    if ($_lines[$_i].StartsWith("#"))
    {
        Write-Host -NoNewLine $("`n[$_i]PS> ")
        Write-Host -NoNewLine -Foreground $CommentColor $($($_Lines[$_i]) + "  ")
        continue
    }else
    {
        # Put the current command in the Window Title along with the demo duration
        $_Duration = [DateTime]::Now - $_StartTime
        $Host.UI.RawUI.WindowTitle = "[{0}m, {1}s]    {2}" -f [int]$_Duration.TotalMinutes, [int]$_Duration.Seconds, $($_Lines[$_i])
        Write-Host -NoNewLine $("`n[$_i]PS> ")
        $_SimulatedLine = $($_Lines[$_i]) + "  "
        for ($_j = 0; $_j -lt $_SimulatedLine.Length; $_j++)
        {
           Write-Host -NoNewLine $_SimulatedLine[$_j]
           if ($_PretendTyping)
           {
               if ([System.Console]::KeyAvailable)
               { 
                   $_PretendTyping = $False
               }
               else
               {
                   Start-Sleep -milliseconds $(10 + $_Random.Next($_InterkeyPause))
               }
           }
        } # For $_j     
        $_PretendTyping = $true    
    } # else


    $_OldColor = $host.UI.RawUI.ForeGroundColor
    $host.UI.RawUI.ForeGroundColor = $MetaCommandColor
    $_input=[System.Console]::ReadLine().TrimStart()
    $host.UI.RawUI.ForeGroundColor = $_OldColor


    switch ($_input)
    {
################ HELP with DEMO
      "?"  
          {
            Write-Host -ForeGroundColor $CommentColor @"
Running demo: $file
. 
(!)  Suspend            (#x) Goto Command #x        (b) Backup  (d) Dump demo
(fx) Find cmds using X  (px) Typing Pause Interval  (q) Quit    (s) Skip 
(t)  Timecheck          (?)  Help                      
. 
NOTE 1: Any key cancels "Pretend typing" for that line.  Use <SPACE> unless you
        want to run a one of these meta-commands.  
.
NOTE 2: After cmd output, enter <CR> to move to the next line in the demo.
        This avoids the audience getting distracted by the next command 
        as you explain what happened with this command.
.
NOTE 3: The line to be run is displayed in the Window Title BEFORE it is typed.
        This lets you know what to explain as it is typing.
"@
            $_i -= 1
          }


      #################### PAUSE
      {$_.StartsWith("p")} 
          {
               $_InterkeyPause = [int]$_.substring(1)
               $_i -= 1
          }


      ####################  Backup
      "b" { 
                if($_i -gt 0)
                {
                    $_i --

                    while (($_i -gt 0) -and ($_lines[$($_i)].StartsWith("#")))
                    {   $_i -= 1
                    }
                }

                $_i --
                $_PretendTyping = $false
          }


      ####################  QUIT
      "q"  
          {
            Write-Host -ForeGroundColor $CommentColor "<Quit demo>"
            return          
          }


      ####################  SKIP
      "s"
          {
            Write-Host -ForeGroundColor $CommentColor "<Skipping Cmd>"
          }


      ####################  DUMP the DEMO
      "d"
          {
            for ($_ni = 0; $_ni -lt $_lines.Count; $_ni++)
            {
               if ($_i -eq $_ni)
               {   Write-Host -ForeGroundColor $MetaCommandColor ("*" * 80)
               } 
               Write-Host -ForeGroundColor $CommentColor ("[{0,2}] {1}" -f $_ni, $_lines[$_ni])
            }
            $_i -= 1
          }


      ####################  TIMECHECK
      "t"  
          {
             $_Duration = [DateTime]::Now - $_StartTime
             Write-Host -ForeGroundColor $CommentColor $(
                "Demo has run {0} Minutes and {1} Seconds`nYou are at line {2} of {3} " -f 
                    [int]$_Duration.TotalMinutes, 
                    [int]$_Duration.Seconds,
                    $_i,
                    ($_lines.Count - 1)
             )
             $_i -= 1
          }


      ####################  FIND commands in Demo
      {$_.StartsWith("f")}
          {
            for ($_ni = 0; $_ni -lt $_lines.Count; $_ni++)
            {
               if ($_lines[$_ni] -match $_.SubString(1))
               {
                  Write-Host -ForeGroundColor $CommentColor ("[{0,2}] {1}" -f $_ni, $_lines[$_ni])
               }
            }
            $_i -= 1
          }


      ####################  SUSPEND
      {$_.StartsWith("!")}
          {
             if ($_.Length -eq 1)
             {
                 Write-Host -ForeGroundColor $CommentColor "<Suspended demo - type 'Exit' to resume>"
                 function Prompt {"[Demo Suspended]`nPS>"}
                 $host.EnterNestedPrompt()
             }else
             {
                 trap [System.Exception] {Write-Error $_;continue;}
                 Invoke-Expression $(".{" + $_.SubString(1) + "}| out-host")
             }
             $_i -= 1
          }


      ####################  GO TO
      {$_.StartsWith("#")}  
          {
             $_i = [int]($_.SubString(1)) - 1
             continue
          }


      ####################  EXECUTE
      default
          {
             trap [System.Exception] {Write-Error $_;continue;}
             Invoke-Expression $(".{" + $_lines[$_i] + "}| out-host")
             $_Duration = [DateTime]::Now - $_StartTime
             $Host.UI.RawUI.WindowTitle = "[{0}m, {1}s]    {2}" -f [int]$_Duration.TotalMinutes, [int]$_Duration.Seconds, $($_Lines[$_i])
             [System.Console]::ReadLine()
          }
    } # Switch
  } # for
  $_Duration = [DateTime]::Now - $_StartTime
  Write-Host -ForeGroundColor $CommentColor $("<Demo Complete {0} Minutes and {1} Seconds>" -f [int]$_Duration.TotalMinutes, [int]$_Duration.Seconds)
  Write-Host -ForeGroundColor $CommentColor $([DateTime]::now)
} # function

