// detailsRoutes.js
const express = require('express');
const detailsController = require('../controllers/detailsController');

const router = express.Router();

router.get('/:id', detailsController.getDetails);

module.exports = router;
