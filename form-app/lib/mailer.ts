import nodemailer from "nodemailer";

const APP_URL = process.env.APP_URL || "http://localhost:3000";

// cache transporter
let transporter: nodemailer.Transporter | null = null;

async function getTransporter() {
  if (transporter) return transporter;

  const testAccount = await nodemailer.createTestAccount();

  transporter = nodemailer.createTransport({
    host: "smtp.ethereal.email",
    port: 587,
    secure: false,
    auth: {
      user: testAccount.user,
      pass: testAccount.pass,
    },
  });

  console.log("Ethereal account created:");
  console.log("USER:", testAccount.user);
  console.log("PASS:", testAccount.pass);

  return transporter;
}

export async function sendInviteEmail(email: string, token: string) {
  const transport = await getTransporter();

  const inviteUrl = `${APP_URL}/api/form/${token}`;

  const info = await transport.sendMail({
    from: '"Feedback App" <no-reply@test.com>',
    to: email,
    subject: "Feedback Form Invitation",
    text: `Open form: ${inviteUrl}`,
    html: `
      <h3>Feedback Form</h3>
      <p>Click below to open:</p>
      <a href="${inviteUrl}">${inviteUrl}</a>
      <p>This link expires in 48 hours.</p>
    `,
  });

  // VERY IMPORTANT
  console.log("Email Preview URL:");
  console.log(nodemailer.getTestMessageUrl(info));

  // EXTRA (Direct token for easy testing)
  console.log(" Token:", token);
}