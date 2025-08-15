const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const locationController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const StatesAndCities = require("../utils/locationData")

locationController.post("/state/list", (req, res) =>
  sendResponse(res, 200, "Success", {
    success: true,
    message: "States retrieve successfully!",
    States: StatesAndCities,
  })
);

locationController.get("/cities/list", (req, res) =>
  sendResponse(res, 200, "Success", {
    success: true,
    message: "Cities retrieve successfully!",
    Cities: City,
  })
);

module.exports = locationController;
