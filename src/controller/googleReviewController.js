const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const axios = require("axios");

const googleReviewController = express.Router();

googleReviewController.get("/google/list", async (req, res) => {
  try {
    const { data } = await axios.get(
      `https://maps.googleapis.com/maps/api/place/details/json`,
      {
        params: {
          place_id: process.env.GOOGLE_PLACE_ID,
          key: process.env.GOOGLE_API_KEY,
          fields: "reviews",
        },
      }
    );

    console.log("Google Reviews API Response:", data);

    if (data.status === "OK" && data.result && data.result.reviews) {
      sendResponse(res, 200, "Success", {
        success: true,
        message: "Reviews retrieved successfully!",
        Reviews: data.result.reviews,
      });
    } else {
      sendResponse(res, 400, "Failed", {
        success: false,
        message: data.error_message || "No reviews found or invalid request",
        data,
      });
    }
  } catch (error) {
    console.error("Error fetching Google Reviews:", error);
    sendResponse(res, 500, "Failed", {
      success: false,
      message: "Failed to fetch Google Reviews",
      error: error.message,
    });
  }
});

module.exports = googleReviewController;
