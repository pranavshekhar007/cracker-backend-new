exports.sendResponse = (res, code, message, data) => {
    if (data) {
      data = data;
    }
    return res.status(code).json(data);
  };
exports.generateOTP = () => {
    const min = 1000;
    const max = 9999;
    return Math.floor(Math.random() * (max - min + 1)) + min;
};