const displayName = _spPageContextInfo.userDisplayName;

const myArray = displayName.split(",");

// Get the current date and time

const currentDate = new Date();

// Get the current hour from the current date

const currentHour = currentDate.getHours();

// Check the current hour and choose a greeting accordingly

let greeting;

if (currentHour < 12) {

greeting = `Guten Morgen, ${myArray[1]} ${myArray[0]}!`;

} else if (currentHour < 18) {

greeting = `Guten Tag, ${myArray[1]} ${myArray[0]}!`;

} else {

greeting = `Guten Abend, ${myArray[1]} ${myArray[0]}!`;

}

// Display the greeting in the page

document.getElementById("greeting").innerText = greeting