// toggleStatusRoutes.js
const express = require('express');
const toggleStatusController = require('../controllers/toggleStatusController');

const router = express.Router();

router.post('/:id', toggleStatusController.toggleStatus);

module.exports = router;
