#####################################################################################
#####################################################################################
## Front End
#####################################################################################
#####################################################################################

$CSSVAR = @"

<style>

@media print {

    

    body {
        
        transform: scale(0.92); 
    
    
    }
    
}


@page { size: auto;  margin: 0mm; }


.HEADER {

    text-align: center;

}

h1 {
    
    top:10px;
    left: 50px;
    position: relative


}

h3 {
    
    float:right;
    margin-right:80px;
    margin-top:-20px;
    color: red;


}

h2,h4 {

    
    left: 50px;
    position: relative



}

body {

    margin: 2px;


}

table {

    border-collapse: collapse;
    margin: 25px 0;
    font-size: 0.9em;
    font-family: sans-serif;
    min-width: 400px;
    box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
    
    position: relative

}

th {

    background-color: #004498;
    color: #ffffff;
    text-align: left;

}



td,th {
    padding: 5px 5px;
    
}

tbody tr {
    border-bottom: 1px solid #dddddd;
}

tbody td:nth-child(1) {
    width: 2%;
   
}

tbody td:nth-child(2) {
    width: 40%;
   
}
tbody td:nth-child(3) {
    width: 2%;
   
}
tbody td:nth-child(4) {
    width: 25%;
   
}
tbody td:nth-child(5) {
    width: 25%;
   
}

tbody tr:nth-of-type(even) {
    background-color: #f3f3f3;
}

tbody tr:last-of-type {
    border-bottom: 2px solid #004498;
}

tbody tr.active-row {
    font-weight: bold;
    color: #009879;
}


</style>



"@
#####################################################################################
#####################################################################################
## Functions
#####################################################################################
#####################################################################################
function Get-MatchingQuestions {
    param (
        $Course_Id,$Quiz_Id,$assignment_id
    )

# get matches and answers
# get Matching Question from here.
$token2 = '*TOKEN*'
$data3 = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($Course_Id)/quizzes/$($Quiz_Id)/questions?per_page=100" -Headers @{"Authorization"="Bearer "+$token2}
$data4 = $data3 | Where-Object {$_.question_type -eq "matching_question"}

$data5  = $data4 | Select-Object question_text -ExpandProperty answers
$data6  = $data4 | Select-Object -ExpandProperty matches

$MatchingQuestionArray = $data4 | Select-Object id



$token2 = '*TOKEN*'
$data2 = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($Course_Id)/assignments/$($assignment_id)/submissions?include[]=submission_history&per_page=100" -Headers @{"Authorization"="Bearer "+$token2}


$data2 = $data2 | Where-Object {($_.workflow_state -ne "unsubmitted") -and ($_.submission_type -ne $null)}

$allstudents = $data2 | Select-Object user_id

$newdata2 = $data2 | Select-Object -ExpandProperty submission_history


$mainDATA = $newdata2 | select  user_id,id -ExpandProperty submission_data



$newdatahello = foreach($x in $MatchingQuestionArray){

    foreach($y in $allstudents){

        $IDtoMatchingObject = [PSCustomObject]@{
    
            user_id = ''
            question_id = ''
    
        
        }
        
        $IDtoMatchingObject.user_id = $y.user_id
        $IDtoMatchingObject.question_id = $x.id
        $IDtoMatchingObject

    }

}


$alltabledata = foreach ($zz in $newdatahello) {

$studentDATA = $mainDATA | where {($_.user_id -eq "$($zz.user_id)") -and ($_.question_id -eq "$($zz.question_id)")}


$AnswersHeader = $studentDATA | Get-member -MemberType 'NoteProperty' | Select-Object 'Name' | where { $_.Name -like 'answer_*'}

$AnswersHeader | ForEach-Object {

    $text = $_.Name
    $_.Name = $text.replace("answer_","")

}


$TableObject = foreach($xy in $AnswersHeader){



    $StudentAnswerObject = [PSCustomObject]@{
    
        user_id = ''
        question_id = ''
        question_text = ''
        matching_question = ''
        student_answer = ''
        correct_answer = ''
        points = ''
    
    }
    
    
    
    
    $CanvasUserId = $studentDATA | Select-Object -ExpandProperty user_id
    $CanvasQuestionId = $studentDATA | Select-Object -ExpandProperty question_id
    $CanvasPoints = $studentDATA | Select-Object -ExpandProperty points
    $StudentAnswerObject.user_id =  $CanvasUserId
    $StudentAnswerObject.question_id = $CanvasQuestionId
    
    

    
    $x = $xy.Name
    
    $something = $studentDATA | Select-Object "answer_$($x)"
    $studentKeyMatchId = $something."answer_$($x)"

    $studentAnswerKey = $data5 | Select-Object * | Where-Object {$_.match_id -eq "$($studentKeyMatchId)"}
    
    
    $studentAnswerKeyMATCH = $data6 | Select-Object * | Where-Object {$_.match_id -eq "$($studentKeyMatchId)"}

    $StudentAnswerRight = $studentAnswerKeyMATCH.text

    

    
    $AnswerKey = $data5 | Select-Object * | Where-Object {$_.id -eq "$($x)"}
    $StudentAnswerObject.matching_question = $AnswerKey.text
    $AnswerRight = $AnswerKey.right

    

    $AnswerKeyMatchId = $AnswerKey.match_id
    


    $StudentAnswerObject.question_text = $AnswerKey.question_text

    if($studentKeyMatchId -ne $AnswerKeyMatchId){

        $StudentAnswerObject.points = "0.00"
        
        $StudentAnswerObject.student_answer = $StudentAnswerRight 


        $StudentAnswerObject.correct_answer = $AnswerRight



    }else{
        $StudentAnswerObject.points = "1.00"
        $StudentAnswerObject.student_answer = $StudentAnswerRight 

        
    }




    $StudentAnswerObject





}


$TableObject  



}




return $alltabledata

    
    
}


function Get-StudentInfoPowercampus {
    param (
        $id
    )
    
    $SQLData = Invoke-Sqlcmd -ServerInstance '*SERVER*' -Username '*USERNAME*' -Password '*PASSWORD*' -Database "*DATABASE*" -Query " 



    select PEOPLE_ID,FIRST_NAME,MIDDLE_NAME,LAST_NAME from PEOPLE where PEOPLE_ID = '$($id)'

    "

    return $SQLData



}

function Check-CanvasCourseQuiz {

    param (
        $CourseId,$QuizID
    )

    
    try{
        #GET /api/v1/courses/:course_id/quizzes/:id

        $token2 = '*TOKEN*'
        $data = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($CourseId)/quizzes/$($QuizID)" -Headers @{"Authorization"="Bearer "+$token2}
        $QuizName = $data.title
        

        return $QuizName
    }
    catch{

        return $false
    }
    
    
    
}

function Check-CanvasCourseExist {

    param (
        $CourseId
    )

    
    try{

        $token2 = '*TOKEN*'
        $data = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($CourseId)" -Headers @{"Authorization"="Bearer "+$token2}
        $CourseName = $data.course_code
        
        return $CourseName
    }
    catch{

        return $false
    }
    
    
    
}

#####################################################################################
#####################################################################################
## Start of Script - MAIN
#####################################################################################
#####################################################################################


$emailinput = Read-Host "Enter Teacher Email"


$CheckingCourse = $false

while($CheckingCourse -eq $false){

Write-Host "Enter Canvas Course ID"

$CourseID = Read-Host " "

$Dacheck = Check-CanvasCourseExist -CourseId "$($CourseID)"

if($Dacheck){
    write-host "Course $($Dacheck) was selected" -ForegroundColor Green
    $CheckingCourse = $true

}
else{

    Write-Host "Sorry Canvas Course doesn't Exist. Try Again." -ForegroundColor Red

}

}




$CheckingCourseQuiz = $false

while($CheckingCourseQuiz -eq $false){

Write-Host "Enter Canvas Quiz ID"

$quizID = Read-Host " "

$Dacheck2 = Check-CanvasCourseQuiz -CourseId "$($CourseID)" -QuizID "$($quizID)"

if($Dacheck2){

    write-host "Quiz $($Dacheck2) was selected" -ForegroundColor Green
    $CheckingCourseQuiz = $true

}
else{

    Write-Host "Sorry Canvas Quiz doesn't Exist. Try Again." -ForegroundColor Red

}

}


$FileName = "$($Dacheck)_$($Dacheck2)"


Write-host "  "
Write-host "  "
Write-host "Working..."




$token2 = '*TOKEN*'
$CanvasStudent = (Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($CourseID)/enrollments?per_page=100" -Headers @{"Authorization"="Bearer "+$token2}) `
| Where-Object {($_.role -eq "StudentEnrollment")} | Select sis_user_id,user_id

$CanvasStudent = $CanvasStudent | Where-Object { $_.sis_user_id -ne $null}




$CanvasStudent | ForEach-Object {

    $_ | Add-Member -MemberType NoteProperty "FirstName" -Value " "
    $_ | Add-Member -MemberType NoteProperty "LastName" -Value " "
    $_ | Add-Member -MemberType NoteProperty "MiddleName" -Value " "

    $x_data = Get-StudentInfoPowercampus -id $_.sis_user_id

    $_.FirstName = $x_data.FIRST_NAME
    $_.LastName = $x_data.LAST_NAME
    $_.MiddleName = $x_data.MIDDLE_NAME



    


    


}



$token2 = '*TOKEN*'
$data = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($CourseID)/quizzes/$($quizID)" -Headers @{"Authorization"="Bearer "+$token2}

$QuizName = $data | select -ExpandProperty title

$assignment_id = $data | select -ExpandProperty assignment_id


# Can get graded quiz here.
$token2 = '*TOKEN*'
$data2 = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($CourseID)/assignments/$($assignment_id)/submissions?include[]=submission_history&per_page=100" -Headers @{"Authorization"="Bearer "+$token2}


$data2 = $data2 | Where-Object {($_.workflow_state -ne "unsubmitted") -and ($_.submission_type -ne $null)}
$newdata2 = $data2 | Select-Object -ExpandProperty submission_history


$mainDATA = $newdata2 | Select-Object  user_id,id,entered_grade -ExpandProperty submission_data #| Export-Csv .\data.csv -NoTypeInformation

$data3 = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/courses/$($CourseID)/quizzes/$($quizID)/questions?per_page=100" -Headers @{"Authorization"="Bearer "+$token2}

$newdataaaaa =  $data3 | Select-Object @{label="question_id"; expression={($_.id)}} -ExpandProperty answers 
$newdataaaaa = $newdataaaaa | Select-Object weight,@{label="answer_id"; expression={($_.id)}},@{label="answer_text"; expression={($_.text)}},question_id
$Question2Answer = $newdataaaaa | Select-Object weight,answer_id,answer_text,question_id | where {($_.weight -ne "0")}


#### Loads PS Script for Join-Object Function ####
. .\Join.ps1


$partialdata = $mainDATA | where-object {$_.correct -clike "partial"}


$partialdata2 = $newdataaaaa | where-object {$_.weight -eq $null}


$newasdfasdf = $mainDATA | Join-Object $newdataaaaa -On question_id,answer_id

$partialdata3 = $partialdata | Join-Object $partialdata2 -On question_id


$newasdfasdf | ForEach-Object {

    $_ | Add-Member -MemberType NoteProperty "Correct_Answer" -Value " "

    if($_.correct -eq $false){

        $questionid = $_.question_id

        $Question2AnswerX = $Question2Answer | select  answer_text,question_id | where {$_.question_id -eq "$($questionid)"}
        $Question2AnswerX = $Question2AnswerX | select -ExpandProperty answer_text
        

        $_.Correct_Answer = $Question2AnswerX



    }


}



$ids = $newasdfasdf | select id | sort * -Unique 

$newstudentDATA = foreach($x in $ids){

    $token2 = '*TOKEN*'
    $data = Invoke-RestMethod -Method Get -Uri "https://*SCHOOLINSTANCE*.instructure.com/api/v1/quiz_submissions/$($x.id)/questions" -Headers @{"Authorization"="Bearer "+$token2}
    
    
    $new = $data | Select-Object -ExpandProperty quiz_submission_questions 
    
    $new | ForEach-Object {
    
        $_.question_text = $_.question_text -replace '<[^>]+>',''
    
    
    }
    
    $newHTML2 = $new | Select-Object question_name,question_text,correct,@{label="question_id"; expression={($_.id)}}
    $newHTML2



}



$finalformdata = $newasdfasdf | Join-Object $newstudentDATA -On question_id

$MatchingQuestionData = Get-MatchingQuestions -Course_Id $CourseID -Quiz_Id $quizID  -assignment_id $assignment_id


foreach ($xx in $CanvasStudent){






    $DATALOOP = $finalformdata | Where-Object user_id -EQ "$($xx.user_id)" |
    Select-Object @{label="Question Number"; expression={[int]($_.question_name -replace "Question ","$1")}},question_text,points,@{label="Student Answer"; expression={($_.answer_text)}},@{label="Correct Answer"; expression={($_.Correct_Answer)}} | Sort-Object  -Unique *



    $total_grade = $finalformdata | Where-Object user_id -EQ "$($xx.user_id)" | Select-Object "entered_grade" | Sort-Object -Unique *
    
    

    $DATALOOP | Sort-Object "Question Number" | ConvertTo-Html -Head "<h1 class='HEADER'>$($xx.FirstName) $($xx.LastName) - Quiz Results</h1><h1 class='HEADER'>$($QuizName)</h1>",$CSSVAR -Precontent "<H4>Quiz Grade: $($total_grade.entered_grade)</H4>" | Out-File ".\$($xx.sis_user_id)_$($xx.FirstName)_$($xx.LastName).html"
    
    # matching Question Logic
    $testgroup = $MatchingQuestionData | Where-Object user_id -EQ "$($xx.user_id)" | Group-Object -Property question_id

    foreach($G in $testgroup){

        $allg = $G | select -ExpandProperty Group 
        $allg = $allg | select -ExpandProperty question_text | sort -Unique
        $G.Group | Select-Object question_id,matching_question,points,student_answer,correct_answer | Sort-Object points -Descending | ConvertTo-Html -Head "<h2>$($allg)</h2>",$CSSVAR | Out-File ".\$($xx.sis_user_id)_$($xx.FirstName)_$($xx.LastName).html" -Append

    }


}


#####################################################################################
#####################################################################################
# Zip and Send Email
#####################################################################################
#####################################################################################

New-Item -ItemType Directory -Name "$($FileName)" -Force

Get-Item .\*.html | Copy-Item -Destination ".\$($FileName)"

Compress-Archive -Path ".\$($FileName)" -DestinationPath ".\$($FileName).zip"

$From = "*FROM EMAIL*"
$To = $emailinput
$Subject = "Quiz Results $($FileName)"
$Body = "Hello Human, 

See attached Zip folder. Use Firefox to print files, better results.

This is an automated task please do not reply to this email.

"
$SMTPServer = "*SMTPSERVER*"
$SMTPPort = "*PORT*"
$pwd = "*PWD*"
$username = "*USERNAME*"
$securepwd = ConvertTo-SecureString $pwd -AsPlainText -Force
$cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $securepwd
$attachedFile = ".\$($FileName).zip"

try{
write-host "Sending..." -ForegroundColor Yellow
Send-MailMessage -From $From -To $To -Subject $Subject -Body $Body -Attachments $attachedFile -SmtpServer $SMTPServer -UseSsl -port $SMTPPort -Credential $cred
Write-Host "Sent." -ForegroundColor Green
}
catch{

Write-Host "Error sending email."


}

Remove-Item .\*.html
Remove-Item ".\$($FileName).zip"
Remove-Item ".\$($FileName)" -Recurse


