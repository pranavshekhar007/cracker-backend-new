const nodemailer = require("nodemailer");

const sendEmail = async (to, subject, html) => {
  // Configure transporter (use your SMTP credentials)
  const transporter = nodemailer.createTransport({
    service: "gmail",
    auth: {
      user: process.env.SMTP_EMAIL,
      pass: process.env.SMTP_PASSWORD, 
    },
  });

  const mailOptions = {
    from: process.env.SMTP_EMAIL,
    to,
    subject,
    html,
  };

  await transporter.sendMail(mailOptions);
};

module.exports = sendEmail;
