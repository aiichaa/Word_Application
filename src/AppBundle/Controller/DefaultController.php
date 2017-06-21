<?php

namespace AppBundle\Controller;

use Sensio\Bundle\FrameworkExtraBundle\Configuration\Route;
use Symfony\Bundle\FrameworkBundle\Controller\Controller;
use Symfony\Component\HttpFoundation\Request;

// Include the BinaryFileResponse and the ResponseHeaderBag
use Symfony\Component\HttpFoundation\BinaryFileResponse;
use Symfony\Component\HttpFoundation\ResponseHeaderBag;

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\TemplateProcessor;
use AppBundle\LibreOfficeWrapper\LibreOffice;


class DefaultController extends Controller
{
    /**
     * @Route("/", name="homepage")
     */
    public function indexAction(Request $request)
    {

        /*// Create a new Word document
        // Create a new Word document
        $phpWord = new PhpWord();
        // Adding an empty Section to the document...
        $section = $phpWord->addSection();
        // Adding Text element to the Section having font styled by default...
        $section->addText(
            '"Learn from yesterday, live for today, hope for tomorrow. '
            . 'The important thing is not to stop questioning." '
            . '(Albert Einstein)'
        );

        // Saving the document as OOXML file...
        $objWriter = IOFactory::createWriter($phpWord, 'Word2007');

        $filePath = 'hello_world_serverfile.docx';
        // Write file into path
        $objWriter->save($filePath);*/


       ///loading .docx template
        //$template = new TemplateProcessor($this->getParameter('template_word').'/dossier.docx');
        $template = new TemplateProcessor($this->getParameter('template_word').'/file.docx');

        //setting single value to placeholder
        //$template->setValue('simulationId', '3333');
        $template->setValue('variableName', 'testttttttt');

        $arrayOne = array("One", "Two", "Three");
        $arrayTwo = array("Volvo", "BMW", "Toyota");
        $arrayThree = array("1999", "1888", "1777");
        $template->cloneRow('ArrayNameOne', count($arrayOne));
        for($number = 0; $number < count($arrayOne); $number++) {
            $template->setValue('ArrayNameOne#'.($number+1), htmlspecialchars($arrayOne[$number], ENT_COMPAT, 'UTF-8'));
            $template->setValue('ArrayNameTwo#'.($number+1), htmlspecialchars($arrayTwo[$number], ENT_COMPAT, 'UTF-8'));
            $template->setValue('ArrayNameThree#'.($number+1), htmlspecialchars($arrayThree[$number], ENT_COMPAT, 'UTF-8'));
        }

        $template->saveAs($this->getParameter('template_word').'/result.docx');

        //use libre office to convert docx to pdf
        $libreOffice = new LibreOffice();
        $libreOffice->convert_docx_to_pdf($this->getParameter('template_word').'/result.docx',$this->getParameter('template_word'));

        /************convert to pdf*************/
        //Load temp file
        //$phpWord = IOFactory::load($this->getParameter('template_word').'/result.docx');
        //$xmlWriter = IOFactory::createWriter($phpWord , 'PDF');
        //$xmlWriter->save($this->getParameter('template_word').'result.pdf');



        // replace this example code with whatever you need
        return $this->render('default/index.html.twig', [
            'base_dir' => realpath($this->getParameter('kernel.project_dir')).DIRECTORY_SEPARATOR,
        ]);
    }
}
