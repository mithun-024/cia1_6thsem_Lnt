// Q3: Convert JSON string into a JavaScript object using JSON.parse() and display the student's name.

let student = '{"name":"Ravi","age":24}';

let studentObj = JSON.parse(student);

console.log("Student Name:", studentObj.name);