// Q4: Create a JavaScript object containing name, age and city.
// Convert it into a JSON string using JSON.stringify() and display the result.

let student = {
    name: "Rahul",
    age: 23,
    city: "Bangalore"
};

let jsonString = JSON.stringify(student);

console.log("JSON String:", jsonString);