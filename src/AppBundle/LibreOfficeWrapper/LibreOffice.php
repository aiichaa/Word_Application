<?php
/**
 * Created by PhpStorm.
 * User: Aicha
 * Date: 21/06/2017
 * Time: 08:24
 */
namespace AppBundle\LibreOfficeWrapper;

//use OCR\TesseractBundle\TesseractWrapper\Exception\CommandException;
use Symfony\Component\HttpFoundation\File\File;

class LibreOffice
{

    /**
     * Path to LibreOffice binary
     *
     * @var string
     */
    protected $path;


    /**
     * Constructor
     *
     * @param string $path Path to Imagick binary (optional)
     */
    public function __construct($path = 'C:\"Program Files\LibreOffice 5"\program\soffice.exe')
    {
        $this->path = $path;
    }


    /**
     * Get version information
     *
     * @return array
     */
    public function getVersion()
    {
        return $this->execute('-version');
    }

    /**
     * Convert docx to pdf
     *
     * @param string $filename
     *
     */
    public function convert_docx_to_pdf($filename)
    {

        $this->execute(
            sprintf(
                ' -headless -convert-to pdf %s ',
                \escapeshellarg($filename)
            )
        );

    }

    /**
     * Execute command and return output
     *
     * @param string $parameters
     * @return array
     *
     * @throws \RuntimeException
     */
    protected function execute($parameters)
    {
        $command = sprintf(
            '%s %s 2>&1',
            $this->path,
            $parameters
        );

        \exec($command, $output, $returnVar);

        if ($returnVar !== 0) {
            // throw CommandException::factory($command, $output);
            var_dump("erreur");
        }

        var_dump("return val : ". $returnVar);

        return $output;
    }
}