function buildClassroom(course){
    Logger.log("Building Classroom for "+course.moduleName);
    // Extract course details from the input object
    const name = course.moduleName;
    const tutor = course.tutorName;
    const tutorEmail = getTutorEmail(tutor);
    const ownerId = tutorEmail.primaryEmail || "me"; // Teacher's email address or "me"

    try {
    // Create the Google Classroom course
    const course = Classroom.Courses.create({
        name: name,
        descriptionHeading: "Welcome to " + name, // Optional heading for the description
        ownerId: ownerId,
    });

    Logger.log("Created classroom with ID: %s", course.id);
    return course.id; 

    } catch (error) {
    Logger.log("Error creating classroom: %s", error.message);
    return null; 
    }
}


function testBuildClassroom(){
    const testCourseSheetId = "1RYs9q9wnoTyL_EdrdpGIchOlNgOn1pmh6ekwyygaZlw";
    const courses = compileCourses(testCourseSheetId);
    const course = courses[0];
    const classroomId = buildClassroom(course);
}