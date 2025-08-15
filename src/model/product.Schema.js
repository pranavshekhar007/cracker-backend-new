const mongoose = require("mongoose");
const timestamps = require("mongoose-timestamp");

const productSchema = mongoose.Schema({
  // step 1
  name: {
    type: String,
    required: true,
  },
  tags: {
    type: [String],
  },
  productType: {
    type: String,
  },
  tax: {
    type: String,
  },
  categoryId: [
    {
      type: mongoose.Schema.Types.ObjectId,
      ref: "Category",
    },
  ],
  hsnCode: {
    type: Number,
  },
  GTIN: {
    type: Number,
  },
  specialAppearance: {
    type: [String],
  },

  shortDescription: {
    type: String,
  },

  // step 2
  stockQuantity: {
    type: Number,
  },
  brandId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: "Brand",
  },
  price: {
    type: Number,
  },
  discountedPrice: {
    type: Number,
  },
  numberOfPieces: {
    type: Number,
  },

  description: {
    type: String,
  },

  soundLevel: { type: String, enum: ["Mild", "Medium", "Loud"] },

  lightEffect: { type: String },

  safetyRating: { type: String, enum: ["KidSafe", "Adult Supervision Only"] },

  usageArea: { type: String, enum: ["Outdoor", "Open Space"] },

  duration: { type: Number },

  weightPerBox: { type: Number },

  // step 3
  productHeroImage: {
    type: String,
  },

  productGallery: {
    type: [String],
  },
  productVideo: {
    type: String,
  },

  status: {
    type: Boolean,
    default: true,
  },
});

productSchema.plugin(timestamps);
module.exports = mongoose.model("Product", productSchema);
