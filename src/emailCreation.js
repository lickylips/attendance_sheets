function getActiveUsers() {
    var users = AdminDirectory.Users.list({
      domain: 'learners.glinnationalcollege.ie', // Replace with your domain
      orderBy: 'email'
    }).users;
  
    for (var i = 0; i < users.length; i++) {
      var user = users[i];
      Logger.log(user.emails);
    }
  }