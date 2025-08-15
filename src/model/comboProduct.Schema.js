const mongoose = require('mongoose');
const timestamps = require("mongoose-timestamp");

const comboProductSchema = new mongoose.Schema(
  {
    // Step 1
    name: {
      type: String,
    },
    productId: [
      {
        product: {
          type: mongoose.Schema.Types.ObjectId,
          ref: 'Product',
        },
        quantity: {
          type: Number,
        },
      }
    ],

    maxComboLimit: {
      type: Number,
    },

    pricing: {
      actualPrice: {
        type: Number,
      },
      offerPrice: {
        type: Number,
      },
      comboPrice: {
        type: Number,
      }
    },
    
    gtin: {
      type: Number,
    },
    shortDescription: {
      type: String,
    },

    // step 2
    stockQuantity: {
      type: String,
    },
  
    longDescription: {
      type: String,
    },

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
      default:true
    },
  },
);

comboProductSchema.plugin(timestamps);
module.exports = mongoose.model('ComboProduct', comboProductSchema);
