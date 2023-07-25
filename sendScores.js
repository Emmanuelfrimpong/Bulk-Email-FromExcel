const xlsx = require('xlsx');
const nodemailer = require('nodemailer');

// Function to send an email
async function sendEmail(senderEmail, senderPassword, receiverEmail, score) {
    if (!receiverEmail) {
        console.error('Error sending email: No recipient email address provided.');
        return;
    }

    if (!score || isNaN(score)) {
        console.error(`Error sending email to ${receiverEmail}: Invalid or missing score.`);
        return;
    }
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: senderEmail,
            pass: senderPassword,
        },
    });

    const mailOptions = {
        from: senderEmail,
        to: receiverEmail,
        subject: 'Your Score',
        text: `Dear recipient,\n\nYour score is: ${score}\n\nBest regards,\nYour Organization`,
    };

    try {
        await transporter.sendMail(mailOptions);
        console.log(`Email sent to ${receiverEmail} with score: ${score}`);
    } catch (error) {
        console.error(`Error sending email to ${receiverEmail}: ${error}`);
    }
}

// Replace these variables with your email and password
const senderEmail = 'emmanuelfrimpong07@gmail.com';
const senderPassword = 'xjhxlfhmcawhuync';

// Read the Excel file
const excelFilePath = 'test_file.xlsx';
const workbook = xlsx.readFile(excelFilePath);
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
const data = xlsx.utils.sheet_to_json(worksheet);

// Iterate through each row and send the email
data.forEach((row) => {
    //print type of row
    console.log(typeof row);
    //convert row to map
    const map = new Map(Object.entries(row));
    //print map
    console.log(map);
    const email = map.get('Email');
    const score = map.get('Score');
    sendEmail(senderEmail, senderPassword, email, score);
});
