const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const chatSchema = mongoose.Schema({
  message: {
    type: String,
  },
  image: {
    type: String,
  },
  userId: {
    type: String,
  },
  userType: {
    type: String,
    enum: ["User", "Vender", "Driver", "Admin"],
  },
  isRead: {
    type: Boolean,
    default: false,
  },
  ticketId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "Ticket",
    required: true,
  },
});

chatSchema.plugin(timestamps);
module.exports = mongoose.model("Chat", chatSchema);
