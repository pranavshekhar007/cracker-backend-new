const Notification = require("../model/notification.Schema");
const admin = require("firebase-admin");
require("dotenv").config(); // Load .env variables

const serviceAccount = {
  type: "service_account",
  project_id: process.env.PROJECT_ID,
  private_key_id: process.env.PRIVATE_KEY_ID,
  private_key: process.env.PRIVATE_KEY?.replace(/\\n/g, '\n'),
  client_email: process.env.CLIENT_EMAIL,
  client_id: process.env.CLIENT_ID,
  auth_uri: process.env.AUTH_URI,
  token_uri: process.env.TOKEN_URI,
  auth_provider_x509_cert_url: process.env.AUTH_PROVIDER_X509_CERT_URL,
  client_x509_cert_url: process.env.CLIENT_X509_CERT_URL,
};

if (!admin.apps.length) {
  admin.initializeApp({
    credential: admin.credential.cert(serviceAccount),
  });
}

exports.sendNotification = async (data) => {
  try {
    const notificationCreated = await Notification.create(data);
    
    const message = {
      notification: {
        title: data?.title || "Default Title",
        body: data?.subTitle || "Default Body",
        image: data?.icon || "Default Body",
      },
      token: "YOUR_FCM_DEVICE_TOKEN_HERE", // Replace or pass dynamically
    };

    const response = await admin.messaging().send(message);
    return notificationCreated;
  } catch (error) {
    console.error("Error creating notification:", error);
    throw error;
  }
};
