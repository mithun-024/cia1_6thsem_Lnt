// Q6: Write JavaScript code using fetch() to retrieve user details from:
// https://jsonplaceholder.typicode.com/users/1
// Display the user's name in the console.

fetch("https://jsonplaceholder.typicode.com/users/1")
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        console.log("User Name:", data.name);
    })
    .catch(error => {
        console.error("Error occurred:", error.message);
    });
