/**
 * Bangladesh Police Recruitment Form - Google Sheets Web App Script
 * 
 * এই কোডটি আপনার গুগল শিটে স্বয়ংক্রিয়ভাবে স্টুডেন্টদের ফর্ম সাবমিশনের ডাটা জমা করবে।
 * প্রতিটি কলাম টেলিটক অফিশিয়াল পুলিশ আবেদনের ক্রমানুসারে সাজানো হয়েছে যাতে আপনি 
 * সহজেই এক ক্লিকে কপি করে টেলিটকে সাবমিট করতে পারেন।
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // কনকারেন্ট সাবমিশন হ্যান্ডেল করার জন্য লক রাখা হয়েছে
  lock.tryLock(15000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // হেডার লিস্ট (টেলিটক পুলিশ আবেদনের ক্রমানুসারে)
    var headers = [
      "Timestamp (সময়)",
      "Applicant Name (নাম)",
      "Father's Name (পিতার নাম)",
      "Mother's Name (মাতার নাম)",
      "Date of Birth (জন্ম তারিখ)",
      "Gender (লিঙ্গ)",
      "Religion (ধর্ম)",
      "Blood Group (রক্তের গ্রুপ)",
      "Quota (কোটা)",
      "Height Feet (ফুট)",
      "Height Inch (ইঞ্চি)",
      "Total Height (মোট উচ্চতা)",
      "Weight (ওজন কেজি)",
      "Chest Normal (বুক স্বাভাবিক ইঞ্চি)",
      "Chest Expanded (বুক প্রসারিত ইঞ্চি)",
      "Nationality (জাতীয়তা)",
      "Birth Reg No (জন্ম নিবন্ধন নম্বর)",
      "NID No (জাতীয় পরিচয়পত্র নম্বর)",
      "Marital Status (বৈবাহিক অবস্থা)",
      
      // বর্তমান ঠিকানা (Present Address)
      "Present Care of",
      "Present Village/Road",
      "Present District (জেলা)",
      "Present Thana/Upazila (থানা)",
      "Present Post Office (পোস্ট অফিস)",
      "Present Post Code (পোস্ট কোড)",
      
      // স্থায়ী ঠিকানা (Permanent Address)
      "Perm Care of",
      "Perm Village/Road",
      "Perm District (জেলা)",
      "Perm Thana/Upazila (থানা)",
      "Perm Post Office (পোস্ট অফিস)",
      "Perm Post Code (পোস্ট কোড)",
      
      // যোগাযোগ
      "Mobile Number (মোবাইল)",
      "Email Address (ইমেইল)",
      
      // এসএসসি তথ্য
      "SSC Exam",
      "SSC Board",
      "SSC Roll",
      "SSC Reg No",
      "SSC Result Type",
      "SSC GPA",
      "SSC Group",
      "SSC Passing Year",
      
      // এইচএসসি তথ্য (যদি থাকে)
      "HSC Exam",
      "HSC Board",
      "HSC Roll",
      "HSC Reg No",
      "HSC Result",
      "HSC GPA",
      "HSC Group",
      "HSC Passing Year",
      
      // অন্যান্য দক্ষতা
      "Skill Type",
      "Skill Particulars",
      
      // ছবি ও স্বাক্ষর
      "Photo Link/Note",
      "Signature Link/Note"
    ];

    // শিট যদি খালি থাকে তবে হেডার তৈরি ও ফরম্যাট করা হবে
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#054f33");
      headerRange.setFontColor("#ffffff");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // ইনকামিং ডাটা পার্স করা
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter;
      }
    } else if (e.parameter) {
      data = e.parameter;
    }

    var timestamp = Utilities.formatDate(new Date(), "Asia/Dhaka", "dd/MM/yyyy HH:mm:ss");

    // ডাটা রো সাজানো
    var rowData = [
      timestamp,
      data.applicant_name || "",
      data.father_name || "",
      data.mother_name || "",
      data.dob_formatted || (data.dob_day + "-" + data.dob_month + "-" + data.dob_year),
      data.gender || "",
      data.religion || "",
      data.blood_group || "",
      data.quota || "",
      data.height_feet || "",
      data.height_inch || "",
      data.total_height || (data.height_feet + "' " + data.height_inch + '"'),
      data.weight_kg || "",
      data.chest_normal || "",
      data.chest_expanded || "",
      data.nationality || "Bangladeshi",
      data.birth_cert_no || "No",
      data.nid_no || "No",
      data.marital_status || "",

      // Present Address
      data.present_care_of || "",
      data.present_village || "",
      data.present_district || "",
      data.present_thana || "",
      data.present_post_office || "",
      data.present_post_code || "",

      // Permanent Address
      data.perm_care_of || "",
      data.perm_village || "",
      data.perm_district || "",
      data.perm_thana || "",
      data.perm_post_office || "",
      data.perm_post_code || "",

      // Mobile & Email
      "'" + (data.mobile_number || ""), // টেক্সট আকারে সেভ হবে যেন সামনের 0 না কাটে
      data.email_address || "",

      // SSC
      data.ssc_exam || "",
      data.ssc_board || "",
      "'" + (data.ssc_roll || ""),
      "'" + (data.ssc_reg || ""),
      data.ssc_result_type || "",
      data.ssc_gpa || "",
      data.ssc_group || "",
      data.ssc_year || "",

      // HSC
      data.hsc_exam || "",
      data.hsc_board || "",
      data.hsc_roll ? "'" + data.hsc_roll : "",
      data.hsc_reg ? "'" + data.hsc_reg : "",
      data.hsc_result_type || "",
      data.hsc_gpa || "",
      data.hsc_group || "",
      data.hsc_year || "",

      // Skills
      data.skill_type || "",
      data.skill_particulars || "",

      // Photo & Signature
      data.photo_link || "",
      data.signature_link || ""
    ];

    sheet.appendRow(rowData);

    // অটো সাইজ কলাম
    // sheet.autoResizeColumns(1, headers.length);

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success",
      "message": "Data saved successfully to Google Sheet",
      "applicant": data.applicant_name
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error",
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ব্রাউজারে টেস্ট করার জন্য doGet
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    "status": "active",
    "message": "Bangladesh Police Recruitment Form Web App is Active and Ready!"
  })).setMimeType(ContentService.MimeType.JSON);
}
