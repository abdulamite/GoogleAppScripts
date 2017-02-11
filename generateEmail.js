function send(e){
  // e.values[i] will pull information from from speadsheet by index
  // example e.values[0] will pull information from column A.

    var userEmail = e.values[16];

    var vist_date = e.values[1];
    var tour_time = e.values[2];
    var student_num = e.values[6];
    var contact_name = e.values[12];
    var lunch = e.values[19];
    var pay_method = e.values[11];
    var want_present = e.values[5];


    var adm_present = "";
    var message = "";
    var subject = "Group Tour Confirmation";

    if(want_present == "yes")
    {
      adm_present = "Presentation time will be included in the final copy of the itinerary.";
    }
    else
    {
      adm_present = "Not Selected";
    }

    if(lunch == "yes")
    {
     message =
    "Dear " + contact_name + ", "
    +"<br />" +
      "Thank you for requesting a Group Tour of Humboldt State University. You have requested the following appointments on " + vist_date + " for your group of "
      + student_num + " students <br />"
      +"Campus Tour Time: " + tour_time + "<br />"
      +"Admissions Presentation Time: " + adm_present + "<br />"
      +"Because you have expressed an interest in lunch on campus, your payment information has been provided below <br />"
      +"Payment option for Lunch  : " + pay_method + "<br />"
      + "Your request will be processed within 5-7 business days. You will receive an itinerary via email confirming " +
         "your appointments, parking instructions and directions to Humboldt State University.";
    }

    else
    {
     message =
     "Dear " + contact_name + ", "
      +"<br />" +
      "Thank you for requesting a Group Tour of Humboldt State University. You have requested the following appointments on " + vist_date + " for your group of "
      + student_num + " students <br />"
      +"Campus Tour Time: " + tour_time + "<br />"
      +"Admissions Presentation Time: " + adm_present + "<br />"
      + "Your request will be processed within 5-7 business days. You will receive an itinerary via email confirming your appointments, parking instructions and directions to Humboldt State University.";
    }

    MailApp.sendEmail({
    to: userEmail,
    subject: subject,
    htmlBody: message
  });
}
