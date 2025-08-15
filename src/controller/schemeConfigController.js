const express = require("express");
const schemeConfigController = express.Router();
const { sendResponse } = require("../utils/common");
const SchemeConfig = require("../model/schemeConfig.Schema");
require("dotenv").config();

/**
 * Admin: Set or Update global scheme config
 */
schemeConfigController.post("/set", async (req, res) => {
  try {
    const { schemeStartDate, schemeEndDate } = req.body;

    if (!schemeStartDate || !schemeEndDate) {
      return sendResponse(res, 400, "Failed", {
        message: "Both start and end dates are required",
        statusCode: 400,
      });
    }

    let config = await SchemeConfig.findOne();

    if (config) {
      config.schemeStartDate = schemeStartDate;
      config.schemeEndDate = schemeEndDate;
      await config.save();
    } else {
      config = await SchemeConfig.create({ schemeStartDate, schemeEndDate });
    }

    sendResponse(res, 200, "Success", {
      message: "Scheme config set successfully!",
      data: config,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

/**
 * Get current scheme config
 */
schemeConfigController.get("/get", async (req, res) => {
  try {
    const config = await SchemeConfig.findOne();

    if (!config) {
      return sendResponse(res, 404, "Failed", {
        message: "Scheme config not found",
        statusCode: 404,
      });
    }

    sendResponse(res, 200, "Success", {
      message: "Scheme config fetched successfully!",
      data: config,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500,
    });
  }
});

schemeConfigController.put("/update/:id", async (req, res) => {
    try {
      const id = req.params.id;
      const { schemeStartDate, schemeEndDate } = req.body;
  
      const config = await SchemeConfig.findById(id);
      if (!config) {
        return sendResponse(res, 404, "Failed", {
          message: "Scheme config not found",
          statusCode: 404,
        });
      }
  
      if (schemeStartDate) config.schemeStartDate = schemeStartDate;
      if (schemeEndDate) config.schemeEndDate = schemeEndDate;
  
      await config.save();
  
      sendResponse(res, 200, "Success", {
        message: "Scheme config updated successfully!",
        data: config,
        statusCode: 200,
      });
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500,
      });
    }
  });
  
  schemeConfigController.delete("/delete/:id", async (req, res) => {
    try {
      const id = req.params.id;
  
      const config = await SchemeConfig.findById(id);
      if (!config) {
        return sendResponse(res, 404, "Failed", {
          message: "Scheme config not found",
          statusCode: 404,
        });
      }
  
      await SchemeConfig.findByIdAndDelete(id);
  
      sendResponse(res, 200, "Success", {
        message: "Scheme config deleted successfully!",
        statusCode: 200,
      });
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500,
      });
    }
  });
  

module.exports = schemeConfigController;
