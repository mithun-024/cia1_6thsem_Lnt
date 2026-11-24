// Q7: Write JavaScript code using Fetch API to retrieve user details from:
// https://jsonplaceholder.typicode.com/users/3
// Display the user's phone number using console.log().

fetch("https://jsonplaceholder.typicode.com/users/3")
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        console.log("Phone Number:", data.phone);
    })
    .catch(error => {
        console.error("Error occurred:", error.message);
    });
