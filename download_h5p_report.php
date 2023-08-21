<?php

// This file is part of Moodle - http://moodle.org/
//
// Moodle is free software: you can redistribute it and/or modify
// it under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// Moodle is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with Moodle.  If not, see <http://www.gnu.org/licenses/>.

/**
 * This file is used to download all attempts or a single attempt of a user in Excel workbook
 * 
 * This file is provided to the button created in local/h5preportdownloader/lib.php
 * When called, writes all attempts or a single attempt of a user in Excel workbook
 * 
 * @package local_h5pdownloadreport
 * @copyright 2023 Kevin Lee
 * @license MIT license, refer to /LICENSE file
 */


require(__DIR__.'/../../config.php');
require_once("lib.php");
require_once("$CFG->libdir/excellib.class.php");
require_once("attempt_writer_functions.php");

use mod_h5pactivity\local\manager;
use mod_h5pactivity\output\result as output_result;


$userid = optional_param('userid', null, PARAM_INT);
$attemptid = optional_param('attemptid', null, PARAM_INT);
$instanceid = required_param('a', PARAM_INT);
$attemtpt;


list ($course, $cm) = get_course_and_cm_from_instance($instanceid, 'h5pactivity');
require_login($course, true, $cm);
$currentgroup = groups_get_activity_group($cm, true);
$manager = manager::create_from_coursemodule($cm);
$report = $manager->get_report($userid, $attemptid, $currentgroup);
$user = $report->get_user();

ob_start();
ob_end_clean();

$name = (string)(($user->username) ? ('_user' . $user->username) : ('_userid' . $userid));
$filename = "h5p_report". $name . ".xls";
$workbook = new MoodleExcelWorkbook($filename);

$xlsformats = new stdClass();
$xlsformats->head1 = $workbook->add_format(['bold' => 1, 'size' => 11]);
$xlsformats->head2 = $workbook->add_format(['bold' => 1, 'size' => 14]);

if($attemptid != null){
    local_h5preportdownloader_write_attempt($workbook, $xlsformats, $report);
}else if($userid != null){
    local_h5preportdownloader_write_all_attempts_in_workbook($workbook, $xlsformats, $manager, $userid);
}

$workbook->close();