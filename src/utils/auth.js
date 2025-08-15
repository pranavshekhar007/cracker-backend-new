const jwt = require("jsonwebtoken");
const User = require("../model/user.Schema"); // Adjust the path based on your project structure

const auth = async (req, res, next) => {
  try {
    const token = req.header("Authorization")?.replace("Bearer ", "");
    if (!token) {
      return res.status(401).json({ error: "Authentication token is required" });
    }

    const decoded = jwt.verify(token, process.env.JWT_KEY); // Ensure JWT_KEY is set in your .env file

    const user = await User.findOne({  _id: decoded.userId });
    if (!user) {
      return res.status(401).json({ error: "User not found or token is invalid" });
    }

    req.token = token;
    req.user = user;
    next();
  } catch (e) {
    res.status(401).json({ error: "Invalid authentication token" });
  }
};

module.exports = auth;
