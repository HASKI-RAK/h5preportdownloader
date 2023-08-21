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
 * This file contains functions that write the h5p attempts to an Excel workbook
 * 
 * @package local_h5pdownloadreport
 * @copyright 2023 Kevin Lee
 * @license MIT license, refer to /LICENSE file
 */

defined('MOODLE_INTERNAL') || die();


require_once(__DIR__.'/../../config.php');
require_once("$CFG->libdir/excellib.class.php");
use mod_h5pactivity\local\manager;
use mod_h5pactivity\output\result as output_result;

/**
 * writes the header of the attempt to the worksheet
 * 
 * @param int $currentrow current row in the worksheet
 * @param MoodleExcelWorksheet $worksheet the Excel worksheet to write on
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_attempt $attempt the h5p attempt to write
 * @param stdClass $user the user who made the attempt
 */
function local_h5preportdownloader_write_attempt_header(&$currentrow, &$worksheet, $xlsformats, $attempt, $user){
    $completion = $attempt->get_completion() ? "Completed" : "Not Completed";
    $completion = "This attempt is " . $completion;
    $success = $attempt->get_success() ? "Success" : "Fail";
    $success = "Outcome: " . $success;
    $totalscore = "Total Score: " . $attempt->get_rawscore() . " out of " . $attempt->get_maxscore();
    $currentrow = 0;
    $worksheet->write_string($currentrow++, 0, (string)date('r',$attempt->get_timecreated()), $xlsformats->head1);
    $worksheet->write_string($currentrow++, 0, $user->username, $xlsformats->head1);
    $worksheet->write_string($currentrow++, 0, $user->email, $xlsformats->head1);
    $worksheet->write_string($currentrow++, 0, 'User ID: '.$user->id, $xlsformats->head1);
    $worksheet->write_string($currentrow++, 0, $completion, $xlsformats->head1);
    $worksheet->write_string($currentrow++, 0, $success, $xlsformats->head1);
    $worksheet->write_string($currentrow++, 0, $totalscore, $xlsformats->head1);
    
    $currentrow++;
    $worksheet->write_string($currentrow, 0, "Questions:", $xlsformats->head2);
    $currentrow += 2;
}

/**
 * writes result of a multiple choice interaction type to the worksheet
 * 
 * @param int $currentrow current row in the worksheet
 * @param MoodleExcelWorksheet $worksheet the Excel worksheet to write on
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_attempt $attempt the h5p attempt to write
 * @param stdClass $user the user who made the attempt
 */
function local_h5preportdownloader_write_on_worksheet_choice (&$currentrow, &$worksheet, $xlsformats, $result){
    global $PAGE;
    $worksheet->write_string($currentrow++, 0, $result->description, $xlsformats->head2);
    $worksheet->write_string($currentrow, 0, "Choice", $xlsformats->head1);
    $worksheet->write_string($currentrow, 1, "Correct Answer", $xlsformats->head1);
    $worksheet->write_string($currentrow++, 2, "User Answer", $xlsformats->head1);
    
    $outputresult = output_result::create_from_record($result);
    $data = $outputresult->export_for_template($PAGE->get_renderer('core'));
    foreach($data->options as $key => $option){
        $iscorrect = str_contains($result->correctpattern, $key);
        $waschosen = str_contains($result->response, $key);
        $worksheet->write_string($currentrow, 0, $option->description);
        $worksheet->write_string($currentrow, 1, $iscorrect);
        $worksheet->write_string($currentrow++, 2, $waschosen);
    }

    $worksheet->write_string($currentrow++, 0, "Score: " . $result->rawscore . " out of " . $result->maxscore, $xlsformats->head1);
    $currentrow++;
}

/**
 * writes result of a sequencing interaction type to the worksheet
 * 
 * @param int $currentrow current row in the worksheet
 * @param MoodleExcelWorksheet $worksheet the Excel worksheet to write on
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_attempt $attempt the h5p attempt to write
 * @param stdClass $user the user who made the attempt
 */
function local_h5preportdownloader_write_on_worksheet_sequencing (&$currentrow, &$worksheet, &$xlsformats, $result){
    global $PAGE;
    $worksheet->set_column(0, 2, 30);
    $worksheet->write_string($currentrow++, 0, $result->description, $xlsformats->head2);
    $worksheet->write_string($currentrow, 0, "Positions", $xlsformats->head1);
    $worksheet->write_string($currentrow, 1, "User Answer", $xlsformats->head1);
    $worksheet->write_string($currentrow++, 2, "Result", $xlsformats->head1);
    
    $outputresult = output_result::create_from_record($result);
    $data = $outputresult->export_for_template($PAGE->get_renderer('core'));

    foreach($data->options as $key => $option){
        $worksheet->write_string($currentrow, 0, $option->description);
        $worksheet->write_string($currentrow, 1, $option->correctanswer->answer);
        $worksheet->write_string($currentrow++, 2, $option->useranswer->answer);
    }
    $worksheet->write_string($currentrow++, 0, "Score: " . $result->rawscore . " out of " . $result->maxscore, $xlsformats->head1);
}

/**
 * writes result of a true false interaction type to the worksheet
 * 
 * @param int $currentrow current row in the worksheet
 * @param MoodleExcelWorksheet $worksheet the Excel worksheet to write on
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_attempt $attempt the h5p attempt to write
 * @param stdClass $user the user who made the attempt
 */
function local_h5preportdownloader_write_on_worksheet_true_false (&$currentrow, &$worksheet, &$xlsformats, $result){
    global $PAGE;
    $worksheet->write_string($currentrow, 0, $result->description, $xlsformats->head2);
    $worksheet->write_string($currentrow, 1, "Correct Answer", $xlsformats->head1);
    $worksheet->write_string($currentrow++, 2, "User Answer", $xlsformats->head1);
    
    $outputresult = output_result::create_from_record($result);
    $data = $outputresult->export_for_template($PAGE->get_renderer('core'));

    foreach($data->options as $option){
        $worksheet->write_string($currentrow, 0, $option->description);
        if(str_contains($option->correctanswer->answer, "unchecked")){
            $worksheet->write_string($currentrow, 1, "Incorrect Answer");
        }
        else{
            $worksheet->write_string($currentrow, 1, "Correct Answer");
        }
        $worksheet->write_string($currentrow++, 2, $option->useranswer->answer);
    }
    $worksheet->write_string($currentrow++, 0, "Score: " . $result->rawscore . " out of " . $result->maxscore, $xlsformats->head1);
    $currentrow++;
}

/**
 * writes result of a fill in interaction type to the worksheet
 * @param int $currentrow current row in the worksheet
 * @param MoodleExcelWorksheet $worksheet the Excel worksheet to write on
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_attempt $attempt the h5p attempt to write
 * @param stdClass $user the user who made the attempt
 */
function local_h5preportdownloader_write_on_worksheet_fill_in (&$currentrow, &$worksheet, &$xlsformats, $result){
    global $PAGE;
    $worksheet->write_string($currentrow++, 0, $result->description, $xlsformats->head2);
    $worksheet->write_string($currentrow, 0, "Matching Choice", $xlsformats->head1);
    $worksheet->write_string($currentrow, 1, "Correct Answer", $xlsformats->head1);
    $worksheet->write_string($currentrow++, 2, "User Answer", $xlsformats->head1);
    
    $outputresult = output_result::create_from_record($result);
    $data = $outputresult->export_for_template($PAGE->get_renderer('core'));

    foreach($data->options as $option){
        $description = str_replace($option->description);
        $worksheet->write_string($currentrow, 0, $option->description);
        $worksheet->write_string($currentrow, 1, $option->correctanswer->answer);
        $worksheet->write_string($currentrow++, 2, $option->useranswer->answer);
    }
    $worksheet->write_string($currentrow++, 0, "Score: " . $result->rawscore . " out of " . $result->maxscore, $xlsformats->head1);
    $currentrow++;
}

/**
 * writes result of a essay interaction type to the worksheet
 * 
 * @param int $currentrow current row in the worksheet
 * @param MoodleExcelWorksheet $worksheet the Excel worksheet to write on
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_attempt $attempt the h5p attempt to write
 * @param stdClass $user the user who made the attempt
 */
function local_h5preportdownloader_write_on_worksheet_long_fill_in (&$currentrow, &$worksheet, &$xlsformats, $result){
    global $PAGE;
    $worksheet->write_string($currentrow++, 0, $result->description, $xlsformats->head2);
    
    $outputresult = output_result::create_from_record($result);
    $data = $outputresult->export_for_template($PAGE->get_renderer('core'));

    $description = str_replace('<p>', '', $data->description);
    $description = str_replace('</p>', '', $description);
    $description = str_replace('__________', '', $description);
    $worksheet->write_string($currentrow++, 0, $description);
    $answer = str_replace('<br />', '', $data->content);
    $answerasarray = explode('.', $answer);
    foreach($answerasarray as $sentence){
        $worksheet->write_string($currentrow++, 0, $sentence);
    }

    $worksheet->write_string($currentrow++, 0, "Score: " . $result->rawscore . " out of " . $result->maxscore, $xlsformats->head1);
}


/**
 * writes result of a Drag and Drop h5p interaction to the worksheet
 * 
 * @param int $currentrow current row in the worksheet
 * @param MoodleExcelWorksheet $worksheet the Excel worksheet to write on
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_attempt $attempt the h5p attempt to write
 * @param stdClass $user the user who made the attempt
 */
function local_h5preportdownloader_write_on_worksheet_matching (&$currentrow, &$worksheet, &$xlsformats, $result){
    global $PAGE;
    $worksheet->set_column(0, 2, 30);
    $worksheet->write_string($currentrow++, 0, $result->description, $xlsformats->head2);
    $worksheet->write_string($currentrow, 0, "Positions", $xlsformats->head1);
    $worksheet->write_string($currentrow, 1, "Correct Answer", $xlsformats->head1);
    $worksheet->write_string($currentrow++, 2, "User Answer", $xlsformats->head1);
    
    $outputresult = output_result::create_from_record($result);
    $data = $outputresult->export_for_template($PAGE->get_renderer('core'));

    foreach($data->options as $key => $option){
        $worksheet->write_string($currentrow, 0, $option->description);
        $worksheet->write_string($currentrow, 1, $option->correctanswer->answer);
        $useranswer = $option->useranswer->answer;
        $answerstatus = ($option->correctanswer->answer == $useranswer) ? "Correct" : "Incorrect";
        $worksheet->write_string($currentrow++, 2, $useranswer . " : " . $answerstatus);
    }
    $worksheet->write_string($currentrow++, 0, "Score: " . $result->rawscore . " out of " . $result->maxscore, $xlsformats->head1);
}

/**
 * writes the attempt to the worksheet
 * 
 * @param MoodleExcelWorkbook $workbook the Excel workbook to write in
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_report $report the h5p report of an attempt to write
 */
function local_h5preportdownloader_write_attempt (&$workbook, $xlsformat, $report){
    $attempt = $report->get_attempt();
    $user = $report->get_user();
    $userid = $attempt->get_userid();
    $results = $attempt->get_results();
    $worksheet = $workbook->add_worksheet();
    $worksheet->set_column(0, 0, 160);
    $worksheet->set_column(1, 2, 15);
    $currentrow = 0;
    
    local_h5preportdownloader_write_attempt_header($currentrow, $worksheet, $xlsformat, $attempt, $user);
    
    foreach ($results as $result){
        if(empty($result->description)) continue;
        if($result->interactiontype == "choice"){
            local_h5preportdownloader_write_on_worksheet_choice ($currentrow, $worksheet, $xlsformats, $result, $PAGE);
        }
        
        if($result->interactiontype == "fill-in"){
            local_h5preportdownloader_write_on_worksheet_fill_in ($currentrow, $worksheet, $xlsformats, $result, $PAGE);
        }
    
        if($result->interactiontype == "sequencing"){
            local_h5preportdownloader_write_on_worksheet_sequencing ($currentrow, $worksheet, $xlsformats, $result, $PAGE);
        }
    
        if($result->interactiontype == "true-false"){
            local_h5preportdownloader_write_on_worksheet_true_false ($currentrow, $worksheet, $xlsformats, $result, $PAGE);
        }
    
        if($result->interactiontype == "long-fill-in"){
            local_h5preportdownloader_write_on_worksheet_long_fill_in ($currentrow, $worksheet, $xlsformats, $result, $PAGE);
        }
    
        if($result->interactiontype == "matching"){
            local_h5preportdownloader_write_on_worksheet_matching ($currentrow, $worksheet, $xlsformats, $result, $PAGE);
        }
    }
}

/**
 * writes all attempts of a particular h5p activity from one user to a Excel workbook
 * 
 * @param MoodleExcelWorkbook $workbook the Excel workbook to write in
 * @param stdClass $xlsformats the formats to use for the worksheet requires head1 and head2 Members
 * @param mod_h5pactivity_classes_local_manager $manager the h5p manager to use
 * @param int $userid id of the user whose attempts to write
 */
function local_h5preportdownloader_write_all_attempts_in_workbook(&$workbook, $xlsformat, $manager, $userid){
    $attempts = $manager->get_user_attempts($userid);
    foreach($attempts as $attempt){
        $attemptid = $attempt->get_id();
        $report = $manager->get_report($userid, $attemptid, $currentgroup);
        local_h5preportdownloader_write_attempt($workbook, $xlsformat, $report);
    }
}

