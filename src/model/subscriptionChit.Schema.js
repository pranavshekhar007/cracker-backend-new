const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const subscriptionChitSchema = mongoose.Schema({
  userId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "User",
  },
  name:{
    type: String,
  },
  phone: {
    type: Number,
  },
  email:{
    type: String,
  },
  location: {
    type:String,
  },
  totalAmount: {
    type: Number,
  },
  monthlyAmount: {
    type: Number,
  },
  totalMonths: {
    type: Number,
  },
  enrolmentDate: {
    type: Date,
    default: Date.now,
  },
  schemeStartDate: {
    type: Date,
  },
  schemeEndDate: {
    type: Date,
  },
  paidMonths: [
    {
      monthNumber: String,
      monthYear: String, 
      paymentDate: Date,
      paymentSs: String,
      status: {
        type: String,
        enum: ["pending", "approved", "rejected"],
        default: "pending",
      },
    },
  ],
   status: {
    type: String,
    enum: ["pending", "approved", "rejected"],
    default: "pending",
  },
  password: {
    type: String,
  }
});

subscriptionChitSchema.plugin(timestamps);
module.exports = mongoose.model("SubscriptionChit", subscriptionChitSchema);
