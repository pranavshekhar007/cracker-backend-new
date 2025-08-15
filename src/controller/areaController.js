const express = require("express");
const { sendResponse } = require("../utils/common");
const Area = require("../model/area.Schema");
const State = require("../model/state.schema");
const City = require("../model/city.Schema");
const Pincode = require("../model/pincode.Schema");
require("dotenv").config();

const areaController = express.Router();

// Create Area
areaController.post("/create", async (req, res) => {
  try {
    const { name, stateId, cityId, pincodeId, minimumPrice, deliveryCharge } = req.body;

    if (!name || !stateId || !cityId || !pincodeId || minimumPrice === undefined || deliveryCharge === undefined) {
      return sendResponse(res, 400, "Failed", {
        message: "All fields are required",
        statusCode: 400,
      });
    }

    // Auto-increment areaId
    const last = await Area.findOne().sort({ areaId: -1 });
    const nextAreaId = last && Number.isInteger(last.areaId) ? last.areaId + 1 : 1;

    const areaCreated = await Area.create({
      areaId: nextAreaId,
      name: name.trim(),
      stateId,
      cityId,
      pincodeId,
      minimumPrice,
      deliveryCharge,
    });

    // Fetch related data from State, City, Pincode models
    const state = await State.findOne({ stateId: stateId });
    const city = await City.findOne({ cityId: cityId });
    const pincode = await Pincode.findOne({ pincodeId: pincodeId });

    sendResponse(res, 200, "Success", {
      message: "Area created successfully!",
      data: {
        ...areaCreated._doc,
        state,
        city,
        pincode,
      },
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});


// List Areas
areaController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      stateId,
      cityId,
      pincodeId,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;

    const query = {};
    if (searchKey) query.name = { $regex: searchKey, $options: "i" };
    if (stateId) query.stateId = stateId;
    if (cityId) query.cityId = cityId;
    if (pincodeId) query.pincodeId = pincodeId;

    const sortField = sortByField || "areaId";
    const sortOrder = sortByOrder === "desc" ? -1 : 1;

    const areas = await Area.find(query)
      .sort({ [sortField]: sortOrder })
      .limit(parseInt(pageCount))
      .skip((parseInt(pageNo) - 1) * parseInt(pageCount));

    // Fetch state, city, pincode data for each area
    const areasWithDetails = await Promise.all(
      areas.map(async (area) => {
        const state = await State.findOne({ stateId: area.stateId });
        const city = await City.findOne({ cityId: area.cityId });
        const pincode = await Pincode.findOne({ pincodeId: area.pincodeId });

        return {
          ...area._doc,
          state,
          city,
          pincode,
        };
      })
    );

    const totalCount = await Area.countDocuments({});
    const filteredCount = await Area.countDocuments(query);

    sendResponse(res, 200, "Success", {
      message: "Area list retrieved successfully!",
      data: areasWithDetails,
      documentCount: {
        totalCount,
        filteredCount,
      },
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});


// Update Area
areaController.put("/update", async (req, res) => {
  try {
    const { _id } = req.body;
    const area = await Area.findById(_id);
    if (!area) {
      return sendResponse(res, 404, "Failed", {
        message: "Area not found",
        statusCode: 404,
      });
    }

    const updatedArea = await Area.findByIdAndUpdate(_id, req.body, { new: true });
    sendResponse(res, 200, "Success", {
      message: "Area updated successfully!",
      data: updatedArea,
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});

// Delete Area
areaController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const area = await Area.findById(id);
    if (!area) {
      return sendResponse(res, 404, "Failed", {
        message: "Area not found",
        statusCode: 404,
      });
    }

    await Area.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Area deleted successfully!",
      statusCode: 200,
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message,
      statusCode: 500,
    });
  }
});


// Get Areas by pincodeId (for dropdowns)
// areaController.post("/get-by-pincode", async (req, res) => {
//   try {
//     const { pincodeId } = req.body;

//     if (!pincodeId) {
//       return sendResponse(res, 400, "Failed", {
//         message: "pincodeId is required",
//         statusCode: 400,
//       });
//     }

//     const areas = await Area.find({ pincodeId, status: true }).sort({ name: 1 });

//     sendResponse(res, 200, "Success", {
//       message: "Areas fetched by pincodeId successfully!",
//       data: areas,
//       statusCode: 200,
//     });
//   } catch (error) {
//     console.error(error);
//     sendResponse(res, 500, "Failed", {
//       message: error.message || "Internal server error",
//       statusCode: 500,
//     });
//   }
// });


module.exports = areaController;
