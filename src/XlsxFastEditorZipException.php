<?php

namespace alexandrainst\XlsxFastEditor;

/**
 * Errors related to ZIP operations, indicating that the external structure of the XLSX document is damaged or with invalid disc access rights.
 * Returns error codes from https://php.net/ziparchive.open
 */
final class XlsxFastEditorZipException extends XlsxFastEditorException
{
}
