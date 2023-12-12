function calculateSemesterGPA(gradesRange) {

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);
  
  var grades = range.getValues();
  var credits = [1,2,1,3,2,2,3,4]
  var totalGradePoints = 0;
  
  for (var i = 0; i < 8; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var gradePoint = 0;
    
    if (grade !== "") {
      if(grade == "A+"){gradePoint = 4.0;}
      else if (grade == "A"){gradePoint = 4.0;}
      else if (grade == "A-"){gradePoint = 3.7;}
      else if (grade == "B+"){gradePoint = 3.3;}
      else if (grade == "B"){gradePoint = 3.0;}
      else if (grade == "B-"){gradePoint = 2.7;}
      else if (grade == "C+"){gradePoint = 2.3;}
      else if (grade == "C"){gradePoint = 2.0}
      else if (grade == "C-"){gradePoint = 1.7;}
      else if (grade == "E"){gradePoint = 0.0;}
    }
    
    totalGradePoints += gradePoint * credit;
  }
  
  var gpa = totalGradePoints / 18;
  console.log(gpa);
  
  return gpa.toFixed(2);
}

function calculateSemesterGPA2(gradesRange) {

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);
  
  var grades = range.getValues();
  var credits = [2,3,2,2,2,3,4]
  var totalGradePoints = 0;
  
  for (var i = 0; i < 7; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var gradePoint = 0;
    
    if (grade !== "") {
      if(grade == "A+"){gradePoint = 4.0;}
      else if (grade == "A"){gradePoint = 4.0;}
      else if (grade == "A-"){gradePoint = 3.7;}
      else if (grade == "B+"){gradePoint = 3.3;}
      else if (grade == "B"){gradePoint = 3.0;}
      else if (grade == "B-"){gradePoint = 2.7;}
      else if (grade == "C+"){gradePoint = 2.3;}
      else if (grade == "C"){gradePoint = 2.0}
      else if (grade == "C-"){gradePoint = 1.7;}
      else if (grade == "E"){gradePoint = 0.0;}
    }
    
    totalGradePoints += gradePoint * credit;
  }
  
  var gpa = totalGradePoints / 18;
  console.log(gpa);
  
  return gpa.toFixed(2);
}

function fullRepeatmodulecount(agpa,arepeate, acminusrep){

  var sheet = SpreadsheetApp.getActiveSheet();
  var gpa = sheet.getRange(agpa).getValue();
  var repeate = sheet.getRange(arepeate).getValue();
  var cminusrep = sheet.getRange(acminusrep).getValue();

  var b = 0;

  if(gpa < 2.0){
    b = cminusrep;
  }

  var totalrep = (b + repeate);

  return totalrep;
}


//calculate total loss credit in sem 1
function totalCreditloss(gradesRange, agpaa){

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);
  var gpa = sheet.getRange(agpaa).getValue();
  
  var grades = range.getValues();
  var credits = [1,2,1,3,2,2,3,4]

  var totalECredit = 00;
  var totalCCredit = 00;

//calculate total E credit
  for (var i = 0; i < 8; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "E"){number1 = 1;}
    }
    
    totalECredit += number1*credit;
  }


//calculate total C credit
  if(gpa < 2.0){
    for (var i = 0; i < 8; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "C-"){number1 = 1;}
    }
    
    totalCCredit += number1*credit;
  }
  
  totalCredit = totalECredit + totalCCredit;
  return totalCredit;

  }}


  //calculate total loss credit in sem 2
function totalCreditloss2(gradesRange, agpaa){

  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);
  var gpa = sheet.getRange(agpaa).getValue();
  
  var grades = range.getValues();
  var credits = [2, 3, 2, 2, 2, 3, 4]

  var totalECredit = 0;
  var totalCCredit = 0;

//calculate total E credit
  for (var i = 0; i < 7; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "E"){number1 = 1;}
    }
    
    totalECredit = (totalECredit + (number1*credit));
  }
  console.log(totalECredit);


//calculate total C credit
  if(gpa < 2.0){
    for (var i = 0; i < 7; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "C-"){number1 = 1;}
    }
    
    totalCCredit = (totalCCredit + (number1*credit));
  }
  console.log(totalCCredit);
  
  totalCredit = (totalECredit + totalCCredit);
  return totalCredit;

  }}

function Ecreditsem1(gradesRange){
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);

  var grades = range.getValues();
  var credits = [1,2,1,3,2,2,3,4]

  var totalECredit = 0;

  //calculate total E credit
  for (var i = 0; i < 8; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "E"){number1 = 1;}
    }
    
    totalECredit += number1*credit;
  }
  return totalECredit;

}
function Ecreditsem2(gradesRange){
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);

  var grades = range.getValues();
  var credits = [2, 3, 2, 2, 2, 3, 4]

  var totalECredit = 0;

  //calculate total E credit
  for (var i = 0; i < 7; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "E"){number1 = 1;}
    }
    
    totalECredit += number1*credit;
  }
  return totalECredit;

}

function Ccreditsem1(gradesRange, agpaa){
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);
  var gpa = sheet.getRange(agpaa).getValue();

  var grades = range.getValues();
  var credits = [1,2,1,3,2,2,3,4]

  var totalCCredit = 0;

  //calculate total C credit
  if(gpa < 2){
    for (var i = 0; i < 8; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "C-"){number1 = 1;}
    }
    
    totalCCredit += number1*credit;
  }
  
  }
  return totalCCredit;
}

function Ccreditsem2(gradesRange, agpaa){
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = SpreadsheetApp.getActiveSpreadsheet().getRangeByName(gradesRange);
  var gpa = sheet.getRange(agpaa).getValue();

  var grades = range.getValues();
  var credits = [2, 3, 2, 2, 2, 3, 4]

  var totalCCredit = 0;

  //calculate total C credit
  if(gpa < 2.0){
    for (var i = 0; i < 7; i++) {
    var grade = grades[0][i];
    var credit = credits[i];
    var number1 = 0;
    
    if (grade !== "") {
      if(grade == "C-"){number1 = 1;}
    }
    
    totalCCredit = (totalCCredit + (number1*credit));
  }
  
  }
  return totalCCredit;
}
