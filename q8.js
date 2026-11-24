// Q8: Write JavaScript code to retrieve all users from:
// https://jsonplaceholder.typicode.com/users
// Use forEach() to display each user's name.

fetch("https://jsonplaceholder.typicode.com/users")
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(users => {
        console.log("List of Users:");
        users.forEach((user, index) => {
            console.log(`${index + 1}. ${user.name}`);
        });
    })
    .catch(error => {
        console.error("Error occurred:", error.message);
    });
