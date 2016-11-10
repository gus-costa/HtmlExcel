<?php

function Zip($source, $destination)
{
if (!extension_loaded('zip') || !file_exists($source)) {
    return false;
}

$zip = new ZipArchive();
if (!$zip->open($destination, ZIPARCHIVE::CREATE)) {
    return false;
}

if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {
    DEFINE('DS', DIRECTORY_SEPARATOR); //for windows
} else {
    DEFINE('DS', '/'); //for linux
}


$source = str_replace('\\', DS, realpath($source));

if (is_dir($source) === true)
{
    $files = new RecursiveIteratorIterator(new RecursiveDirectoryIterator($source), RecursiveIteratorIterator::SELF_FIRST);
    //echo $source;
    foreach ($files as $file)
    {
        $file = str_replace('\\',DS, $file);
        // Ignore "." and ".." folders
        if( in_array(substr($file, strrpos($file, DS)+1), array('.', '..')) )
            continue;

        $file = realpath($file);

        if (is_dir($file) === true)
        {
            $zip->addEmptyDir(str_replace($source . DS, '', $file . DS));
        }
        else if (is_file($file) === true)
        {
            $zip->addFromString(str_replace($source . DS, '', $file), file_get_contents($file));
        }
        //echo $source;
    }
}
else if (is_file($source) === true)
{
    $zip->addFromString(basename($source), file_get_contents($source));
}

return $zip->close();
}

Zip("xlsm", "test.xlsm");