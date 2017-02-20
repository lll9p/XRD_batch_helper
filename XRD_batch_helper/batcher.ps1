$WORK_DIR = Split-Path -Parent $MyInvocation.MyCommand.Definition
$inp = $WORK_DIR+"\INP\QA.INP"
$result_file = $WORK_DIR+"\result.txt"
$inp_test = $WORK_DIR+"\INP\QT_TMP.INP"
$files = Get-ChildItem -Path $WORK_DIR\*.raw
function get_params($result_path)
{   
    $lines = (Get-Content $result_path)
    $params = [System.Collections.ArrayList]@()
    $values = [System.Collections.ArrayList]@()
    foreach($line in $lines){
        if(($line -match '^r_exp.*(r_wp)  (\d+\.\d+)')){
            $_T = $params.Add($Matches[1])
                $_T = $values.Add($Matches[2])
        }
        if(($line -match 'phase_name "(.*)"'))
        {
           $_T = $params.Add($Matches[1])
        }
        if($line -match 'MVW\(.* (\d+\.\d+)\`\)$')
        {
            $_T = $values.Add($Matches[1])
        }
    }
    return $params,$values
}
$params_got = 0
foreach($file in $files)
{   $file_fullname = $file.FullName
    (Get-Content $inp) -Replace('xdd .*',"xdd `"$file_fullname`"") -join "`r`n"|Set-Content -Path $inp_test
    $execs = '"'+$inp_test.SubString(0,$inp_test.Length-4)+'" "macro FileName { '+$file.FullName.SubString(0,$file.FullName.Length-4)+' }"'
    c:\topas5\tc $execs
    $result_old = $inp_test.SubString(0,$inp_test.Length-4)+'.out'
    $result_new = $file_fullname.SubString(0,$file_fullname.Length-4)+'.out'
    Move-Item $result_old $result_new -force
    Remove-Item  $inp_test
    $params,$values = get_params($result_new)
    $params.Insert(0,"SAMPLE_ID")
    $values.Insert(0,$file.Name.SubString(0,$file.Name.Length-4))
    if($params_got -eq 0){
        $params -join "`t" | Add-Content $result_file
        $params_got = 1
    }
    $values -join "`t" | Add-Content $result_file
}
