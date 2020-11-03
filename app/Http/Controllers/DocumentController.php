<?php

namespace App\Http\Controllers;
use PhpOffice\PhpWord\PhpWord;

class DocumentController extends Controller
{
    public $word;
    public $writer;

    public function __construct()
    {
        $this->word = new PhpWord();
        $this->headerFontStyling();
        $this->attributeFontStyling();
    }

    public function index()
    {
        $availableBlocks = [
            'project',
            'useCases',
            'tasks'
        ];

        $this->frontPage();
        $has_project = in_array('project', $availableBlocks);
        if($has_project) {
            $this->projectParser();
        }
        $this->save();
        return response()->json(['msg' => "Created"], 201);
    }

    public function save()
    {
        $this->writer = \PhpOffice\PhpWord\IOFactory::createWriter($this->word, 'Word2007');
        $this->writer->save('helloWorld.docx');
    }

    public function frontPage()
    {
        // Adding an empty Section to the document...
        $section = $this->addSection();
        $section->addImage(
            'https://images.unsplash.com/photo-1604376120598-e7ecbf34d6bc?ixlib=rb-1.2.1&ixid=eyJhcHBfaWQiOjEyMDd9&auto=format&fit=crop&w=1797&q=80',
            array(
                'width'         => '250',
                'height'        => '250',
                'marginTop'     => -1,
                'marginLeft'    => -1,
                'wrappingStyle' => 'behind'
            )
        );
        $section->addText('Pudding Inc.', 'headerFont');
    }

    public function projectParser()
    {
        // Adding an empty Section to the document...
        $section = $this->addSection();
        // Adding Text element to the Section having font styled by default...
        $section->addText('Projects', 'headerFont');
        $section->addText('Customer: 3M Health Care', 'attributeFont');
        $section->addText('Start Date: 20/2/2020', 'attributeFont');
        $section->addText('End Date: 31/12/2020', 'attributeFont');
        $section->addText('Status: Active', [
            'name' => 'Inter', 'size' => 12, 'color' => '1B2232'
        ]);
        $section->addTextBreak(1);
        $table = $section->addTable();
        for ($r = 1; $r <= 8; $r++) {
            $table->addRow();
            for ($c = 1; $c <= 5; $c++) {
                $table->addCell(1750)->addText(htmlspecialchars("Row {$r}, Cell {$c}"));
            }
        }
    }
    // wrapper
    public function addSection()
    {
        return $this->word->addSection();
    }
    // defining font style
    public function attributeFontStyling()
    {
        $fontStyleName = 'attributeFont';
        $this->word->addFontStyle(
            $fontStyleName,
            array('name' => 'Inter', 'size' => 12, 'color' => '1B2232', 'bold' => false)
        );
    }
    // defining font style
    public function headerFontStyling()
    {
        $fontStyleName = 'headerFont';
        $this->word->addFontStyle(
            $fontStyleName,
            array('name' => 'Inter', 'size' => 14, 'bold' => true)
        );
    }

    public function usingHTML()
    {
        // New Word Document
        $phpWord = $this->word;
        $phpWord->addParagraphStyle('Heading2', array('alignment' => 'center'));

        $section = $phpWord->addSection();
        $html = '<h1>Adding element via HTML</h1>';
        $html .= '<p>Some well-formed HTML snippet needs to be used</p>';
        $html .= '<p>With for example <strong>some<sup>1</sup> <em>inline</em> formatting</strong><sub>1</sub></p>';

        $html .= '<p>A link to <a href="http://phpword.readthedocs.io/" style="text-decoration: underline">Read the docs</a></p>';

        $html .= '<p lang="he-IL" style="text-align: right; direction: rtl">היי, זה פסקה מימין לשמאל</p>';

        $html .= '<p style="margin-top: 240pt;">Unordered (bulleted) list:</p>';
        $html .= '<ul><li>Item 1</li><li>Item 2</li><ul><li>Item 2.1</li><li>Item 2.1</li></ul></ul>';

        $html .= '<p style="margin-top: 240pt;">1.5 line height with first line text indent:</p>';
        $html .= '<p style="text-align: justify; text-indent: 70.9pt; line-height: 150%;">Lorem ipsum dolor sit amet, <strong>consectetur adipiscing elit</strong>, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.</p>';

        $html .= '<h2 style="align: center">centered title</h2>';

        $html .= '<p style="margin-top: 240pt;">Ordered (numbered) list:</p>';
        $html .= '<ol>
                <li><p style="font-weight: bold;">List 1 item 1</p></li>
                <li>List 1 item 2</li>
                <ol>
                    <li>sub list 1</li>
                    <li>sub list 2</li>
                </ol>
                <li>List 1 item 3</li>
            </ol>
            <p style="margin-top: 15px;">A second list, numbering should restart</p>
            <ol>
                <li>List 2 item 1</li>
                <li>List 2 item 2</li>
                <li>
                    <ol>
                        <li>sub list 1</li>
                        <li>sub list 2</li>
                    </ol>
                </li>
                <li>List 2 item 3</li>
                <ol>
                    <li>sub list 1, restarts with a</li>
                    <li>sub list 2</li>
                </ol>
            </ol>';

        $html .= '<p style="margin-top: 240pt;">List with formatted content:</p>';
        $html .= '<ul>
                <li>
                    <span style="font-family: arial,helvetica,sans-serif;">
                        <span style="font-size: 16px;">big list item1</span>
                    </span>
                </li>
                <li>
                    <span style="font-family: arial,helvetica,sans-serif;">
                        <span style="font-size: 10px; font-weight: bold;">list item2 in bold</span>
                    </span>
                </li>
            </ul>';

        $html .= '<p style="margin-top: 240pt;">A table with formatting:</p>';
        $html .= '<table align="center" style="width: 50%; border: 6px #0000FF double;">
                <thead>
                    <tr style="background-color: #FF0000; text-align: center; color: #FFFFFF; font-weight: bold; ">
                        <th style="width: 50pt">header a</th>
                        <th style="width: 50px">header          b</th>
                        <th style="background-color: #FFFF00; border-width: 12px"><span style="background-color: #00FF00;">header c</span></th>
                    </tr>
                </thead>
                <tbody>
                    <tr><td style="border-style: dotted; border-color: #FF0000">1</td><td colspan="2">2</td></tr>
                    <tr><td>This is <b>bold</b> text</td><td></td><td>6</td></tr>
                </tbody>
            </table>';

        $html .= '<p style="margin-top: 240pt;">Table inside another table:</p>';
        $html .= '<table align="center" style="width: 80%; border: 6px #0000FF double;">
    <tr><td>
        <table style="width: 100%; border: 4px #FF0000 dotted;">
            <tr><td>column 1</td><td>column 2</td></tr>
        </table>
    </td></tr>
    <tr><td style="text-align: center;">Cell in parent table</td></tr>
</table>';

        $html .= '<p style="margin-top: 240pt;">The text below is not visible, click on show/hide to reveil it:</p>';
        $html .= '<p style="display: none">This is hidden text</p>';

        \PhpOffice\PhpWord\Shared\Html::addHtml($section, $html, false, false);
// Saving the document as OOXML file...
        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save('helloWorld.docx');
    }
}
