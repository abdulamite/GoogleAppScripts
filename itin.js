function createDoc(e){

  // e.values[i] will pull information from from speadsheet by index
  // example e.values[0] will pull information from column A.
    var subject = "Group Tour Itinerary for " + e.values[14];

    //info from the sheets
    var visit_date = e.values[1];
    var group_name = e.values[14];
    var numStudents = e.values[6];
    var tourTime = e.values[2];
    var payment_method = e.values[11];
    var name =  e.values[12];
    var lunch = e.values[19];
    var presentation = e.values[5];
    var transport = e.values[9];

    //format date options
    var day = new Date(visit_date);
    var options = {
        weekday: "long", year: "numeric", month: "short",
        day: "numeric"
    };

    var long_form_day = day.toLocaleDateString(day,options);



    // Create a new Google Doc using the group name and the visit date
    var doc = DocumentApp.create('Visitor Itinerary--'+group_name + '--' + visit_date);

    doc.getBody().appendParagraph("GROUP/VIP TOUR ITINERARY");
    doc.getBody().appendParagraph(visit_date);

    // Access the body of the document, then add a paragraph.
    doc.getBody().appendParagraph("Hi " + name + ",");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Thank you for setting up a group visit to Humboldt State University on " + long_form_day + ". We look forward to meeting your group of "+ numStudents +" students from "+ group_name +" and informing you about our exceptional educational programs, opportunities for hands on learning, and our many other special programs.");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Here is the itinerary we have set for your visit: ");
    doc.getBody().appendParagraph(long_form_day);
    doc.getBody().appendParagraph("");


    if(lunch == 'yes'){
      doc.getBody().appendParagraph("Lunch available at the Jolly Giant Commons, 'The J' pay with " + payment_method);
      doc.getBody().appendParagraph("We have arranged a special offer for your group to eat in our main cafeteria on the top floor of the Jolly Giant Commons building where your tour ends. This is where our residential students enjoy most of their meals. We can offer an all-you-can-eat lunch for just $5.50 per person (including tax). This lunch includes our hot meal offerings, salad bar, hot soups, ice cream machine, and complete fountain of sodas, juices and coffee drinks. It does not include the sandwich bar or any of the pre-packages foods/drinks for sale in the cafeteria (such as chips, cookies, yogurts or bottled drinks.)");
      doc.getBody().appendParagraph("");
    }

    if(presentation == 'yes')
    {
        doc.getBody().appendParagraph("Admissions Presentation (starts in the Student Business Services Building Lobby)");
        doc.getBody().appendParagraph("");
    }

    doc.getBody().appendParagraph(tourTime);
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Tour of the Campus (starts in the Student Business Services Building Lobby). Please wear comfortable shoes and dress in layers ready for rain, cold or varying weather. There are many stairs and hills on our campus and we'll walk most of them.");
    doc.getBody().appendParagraph("In an effort to respect the privacy and safety of our students, the housing portion of our tours will not include an interior view of our residence hall communities. We do offer full tours of all of our different housing communities during the following campus events: Spring Preview and Fall Admissions Day. You can enjoy a virtual tour of all our communities online 24 hours a day, 365 days a year from the comfort of your own home by visiting:");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("http://www.humboldt.edu/housing/options/");
    doc.getBody().appendParagraph("");

    doc.getBody().appendParagraph("When you're ready you can depart from the first floor of the Jolly Giant Commons. It works well to have your bus or vehicles drive there to pick up the group, rather than having everyone walk back across campus, but this depends on what your driver is comfortable with. The Jolly Giant Commons is on Granite Ave, which is the first right off of L K Wood, one block north of the Sunset Ave freeway exit.");
    doc.getBody().appendParagraph("");

    if(transport == '1 Bus' || transport == '2 or more Buses' )
    {
      doc.getBody().appendParagraph("Directions and Parking (for Buses)");
      doc.getBody().appendParagraph("");
      doc.getBody().appendParagraph("");
      doc.getBody().appendParagraph("Buses should park on Laurel Ave next to the John Van Duzer Theatre. Bus driver needs to stay in the bus at all times when parked on Laurel drive.From the north on HWY 101 (HWY 299 ends at Hwy 101 just north of campus), take the Sunset Ave/Humboldt State University exit, then cross left over the freeway, go right on L K Wood, and take the second left into campus on Harpst Street. At the stop sign at B Street go left, B Street will dead end at Laurel, go left and immediately park along the red curb on the right side. They may unload there and exit straight ahead through the library parking lot, and return to pick up students at the same place. Due to the high volume of traffic, please do not unload in front of the library on Plaza Avenue.");
      doc.getBody().appendParagraph("");
    }
    else{
      doc.getBody().appendParagraph(" ");
      doc.getBody().appendParagraph("Directions and Parking (for cars/vans)");
      doc.getBody().appendParagraph(" ");
      doc.getBody().appendParagraph("From the south on HWY 101 take the 14th Street/Humboldt State University exit, go straight at the stop sign on LK Wood and take the first right into the university on Harpst Street. From the north on HWY 101 (HWY 299 ends at Hwy 101 just north of campus), take the Sunset Ave/Humboldt State University exit, then cross left over the freeway, go right on L K Wood, and take the second left into campus on Harpst Street. Park anywhere in the parking lot along the left side of Harpst Street or in the meters. Then get a guest parking permit at the parking booth in the north east part of the parking lot (or at University Police inside the Student Business Services Building). A parking permit will be waiting for you at the booth. The tour will begin in the lobby of the Student Business Services Building, which is right across the street from the parking lot, to the east.");
      doc.getBody().appendParagraph(" ");
    }

    doc.getBody().appendParagraph("Expectations of your group:");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("For groups of visitors in High School or Middle School, you will need to provide an adult teacher or chaperone for every 20 students in your group.");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("It will be the responsibility of those adults to help keep the students attentive and respectful to the tour guide throughout the tour. Our guides, usually Humboldt State students, have a wealth of information to share about HSU's exceptional academic offering and facilities. With younger groups we lighten up the amount of information, but we still expect groups to be attentive, and stay focused. If at any time participants are disrespectful or excessively disruptive, and their accompanying adult does not correct the situation, the tour guide may conclude the tour. It will then be the responsibility of the adult chaperons to have a backup plan for their group until the next event on their itinerary.");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Again, we look forward to having your group on campus! Please contact me directly if you would like to make adjustments to this itinerary, or if there is anything else you need. Here are a few other links that may be helpful:");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Directions to HSU: http://www.humboldt.edu/humboldt/visit/directions");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Printable Campus Map: http://www.humboldt.edu/humboldt/maps");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Lodging Information: http://www.humboldt.edu/humboldt/visit/travel");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Best Regards,");
    doc.getBody().appendParagraph("");
    doc.getBody().appendParagraph("Jesica");

    // Get the URL of the document.
    var url = doc.getUrl();

    // Append a new string to the "url" variable to use as an email body.
    var message = 'Here is the the link to the generated document: ' + url;


    MailApp.sendEmail({
    to: "jesica.bishop@humboldt.edu",
    subject: subject,
    htmlBody: message
  });
}
