<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\ThuocController;

Route::get('/', [ThuocController::class, 'index']);
Route::post('/add-ma-thuoc-dv/process', [ThuocController::class, 'process']);
Route::post('/add-ma-thuoc-kigui/process', [ThuocController::class, 'processKiGui']);
Route::post('/add-ma-vtyt-dv/process', [ThuocController::class, 'processVTYTDV']);
Route::post('/add-ma-vtyt-kigui/process', [ThuocController::class, 'processVTYTKiGui']);
