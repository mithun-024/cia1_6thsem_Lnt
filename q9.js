// Q9: Write JavaScript code using Fetch API to retrieve user details from:
// https://jsonplaceholder.typicode.com/users/4
// Display the user's company name and website in the console.

fetch("https://jsonplaceholder.typicode.com/users/4")
    .then(response => {
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        return response.json();
    })
    .then(data => {
        console.log("Company Name:", data.company.name);
        console.log("Website:", data.website);
    })
    .catch(error => {
        console.error("Error occurred:", error.message);
    });
