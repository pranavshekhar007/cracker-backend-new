const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");
const { type } = require("os");

const productFaqSchema = mongoose.Schema({
  question: {
    type: String,
  },
  answer: {
    type: String,
  },
  category: {
    type: String,
  },
  productId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "Product",
  },
  venderId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "Vender",
  },
});

productFaqSchema.plugin(timestamps);
module.exports = mongoose.model("ProductFaq", productFaqSchema);
