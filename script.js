let students = [];

function toBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.readAsDataURL(file);
    reader.onload = () => resolve(reader.result);
    reader.onerror = error => reject(error);
  });
}

document.getElementById("admissionForm").addEventListener("submit", async function(e) {
  e.preventDefault();

  const photo = document.getElementById("photo").files[0];
  const sign = document.getElementById("sign").files[0];

  const photo64 = await toBase64(photo);
  const sign64 = await toBase64(sign);

  let student = {
    "DATE OF ADMISSION": doa.value,
    "ADM NO.": admno.value,
    "ONLINE ROLL NO.": rollno.value,
    "NAME OF STUDENT": name.value,
    "DOB": dob.value,
    "FATHER NAME": father.value,
    "MOTHER NAME": mother.value,
    "CLASS": document.getElementById("class").value,
    "ADDRESS": address.value,
    "CAST NAME": cast.value,
    "CATEGORY": category.value,
    "DISTRICT": district.value,
    "PIN CODE": pincode.value,
    "ADHAR NO. STUDENT": adharStudent.value,
    "ADHAR FATHER": adharFather.value,
    "ADHAR MOTHER": adharMother.value,
    "MOBILE NO.": mobile.value,
    "PRIVIOUS SCHOOL": school.value,
    "SUBJECT": subject.value,
    "MEDIUM": medium.value,
    "FORM UPLD": formupld.value,
    "FEES RECEIVED": fees.value,
    "STUDENT PHOTO": photo64,
    "STUDENT SIGNATURE": sign64
  };

  students.push(student);
  alert("Student Added Successfully");
  this.reset();
});

function exportExcel() {
  if (students.length === 0) {
    alert("No Data Available");
    return;
  }

  let wb = XLSX.utils.book_new();
  let classWise = {};

  students.forEach(s => {
    if (!classWise[s.CLASS]) classWise[s.CLASS] = [];
    classWise[s.CLASS].push(s);
  });

  for (let cls in classWise) {
    let ws = XLSX.utils.json_to_sheet(classWise[cls]);
    XLSX.utils.book_append_sheet(wb, ws, cls);
  }

  XLSX.writeFile(wb, "Student_Admission_Data.xlsx");
}
