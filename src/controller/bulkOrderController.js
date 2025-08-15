const express = require("express");
const { sendResponse } = require("../utils/common");
require("dotenv").config();
const BulkOrder = require("../model/bulkOrder.Schema");
const bulkOrderController = express.Router();
require("dotenv").config();
const cloudinary = require("../utils/cloudinary");
const upload = require("../utils/multer");
const { sendNotification } = require("../utils/sendNotification");

bulkOrderController.post(
  "/create",
  upload.single("image"),
  async (req, res) => {
    try {
      let obj;
      if (req.file) {
        let image = await cloudinary.uploader.upload(
          req.file.path,
          function (err, result) {
            if (err) {
              return err;
            } else {
              return result;
            }
          }
        );
        obj = { ...req.body, image: image.url };
      }
      const bulkOrderCreated = await BulkOrder.create(obj);
    //   sendNotification(
    //     {
    //       icon: CategoryCreated?.image,
    //       title: ` has re-uploaded the details`,
    //       subTitle: ` has re-uploaded the details`,
    //       notifyUserId: "Admin",
    //       category: "Driver",
    //       subCategory: "Profile update",
    //       notifyUser: "Admin",
    //     },
    //     req.io
    //   );
      sendResponse(res, 200, "Success", {
        message: "Bulk order created successfully!",
        data: bulkOrderCreated,
        statusCode: 200,
      });
    } catch (error) {
      console.error(error);
      sendResponse(res, 500, "Failed", {
        message: error.message || "Internal server error",
        statusCode: 500,
      });
    }
  }
);

bulkOrderController.post("/list", async (req, res) => {
  try {
    const {
      searchKey = "",
      status,
      pageNo = 1,
      pageCount = 10,
      sortByField,
      sortByOrder,
    } = req.body;
    const query = {};
    if (status) query.status = status;
    if (searchKey) query.name = { $regex: searchKey, $options: "i" };
    const sortField = sortByField || "createdAt";
    const sortOrder = sortByOrder === "asc" ? 1 : -1;
    const sortOption = { [sortField]: sortOrder };
    const bulkOrderList = await BulkOrder.find(query)
      .sort(sortOption)
      .limit(parseInt(pageCount))
      .skip(parseInt(pageNo - 1) * parseInt(pageCount));
    const totalCount = await BulkOrder.countDocuments({});
    const activeCount = await BulkOrder.countDocuments({ status: "orderPlaced" });
    const completedCount = await BulkOrder.countDocuments({ status: "completed" });
    const cancelledCount = await BulkOrder.countDocuments({ status: "cancelled" });
    sendResponse(res, 200, "Success", {
      message: "Bulk order list retrieved successfully!",
      data: bulkOrderList,
      documentCount: {
        totalCount,
        activeCount,
        completedCount,
        cancelledCount
      },
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

bulkOrderController.put("/update", upload.single("image"), async (req, res) => {
  try {
    const id = req.body._id;
    const bulkOrder = await BulkOrder.findById(id);
    if (!bulkOrder) {
      return sendResponse(res, 404, "Failed", {
        message: "Order not found",
        statusCode: 403
      });
    }
    let updatedData = { ...req.body };
    if (req.file) {
      if (bulkOrder.image) {
        const publicId = bulkOrder.image.split("/").pop().split(".")[0];
        await cloudinary.uploader.destroy(publicId, (error, result) => {
          if (error) {
            console.error("Error deleting old image from Cloudinary:", error);
          } else {
            console.log("Old image deleted from Cloudinary:", result);
          }
        });
      }
      const image = await cloudinary.uploader.upload(req.file.path);
      updatedData.image = image.url;
    }
    const updatedBulkOrder = await BulkOrder.findByIdAndUpdate(id, updatedData, {
      new: true, // Return the updated document
    });
    sendResponse(res, 200, "Success", {
      message: "Order updated successfully!",
      data: updatedBulkOrder,
      statusCode: 200
    });
  } catch (error) {
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
      statusCode: 500
    });
  }
});

bulkOrderController.delete("/delete/:id", async (req, res) => {
  try {
    const { id } = req.params;
    const bulkOrder = await BulkOrder.findById(id);
    if (!bulkOrder) {
      return sendResponse(res, 404, "Failed", {
        message: "Order not found",
      });
    }
    const imageUrl = bulkOrder.image;
    if (imageUrl) {
      const publicId = imageUrl.split("/").pop().split(".")[0]; // Extract public ID
      // Delete the image from Cloudinary
      await cloudinary.uploader.destroy(publicId, (error, result) => {
        if (error) {
          console.error("Error deleting image from Cloudinary:", error);
        } else {
          console.log("Cloudinary image deletion result:", result);
        }
      });
    }
    await BulkOrder.findByIdAndDelete(id);
    sendResponse(res, 200, "Success", {
      message: "Order and associated image deleted successfully!",
      statusCode:200
    });
  } catch (error) {
    console.error(error);
    sendResponse(res, 500, "Failed", {
      message: error.message || "Internal server error",
    });
  }
});

module.exports = bulkOrderController;
