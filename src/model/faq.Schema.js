const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const faqSchema = mongoose.Schema({
  question: {
    type: String,
  },
  answer: {
    type: String,
  },
  category :{
    type: String
  }
});

faqSchema.plugin(timestamps);
module.exports = mongoose.model("Faq", faqSchema);
