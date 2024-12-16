const express = require('express');
const router = express.Router();
const inventoryController = require('../controllers/inventory');
const upload = require('../middlewares/upload');

router.post('/add', upload.single('image'), inventoryController.addInventoryWithImage);
router.get('/search', inventoryController.searchInventory);
router.post('/issue', inventoryController.issueInventory);
router.post('/receive', inventoryController.receiveInventory);
router.post('/addNewItem', upload.single('image'), inventoryController.addNewItemWithoutCode);

module.exports = router;