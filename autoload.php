<?php

/**
 * Modified from https://github.com/php-fig/fig-standards/blob/master/accepted/PSR-4-autoloader-examples.md
 */

declare(strict_types=1);

spl_autoload_register(function (string $class): void {

	// project-specific namespace prefix
	$prefix = 'alexandrainst\\XlsxFastEditor\\';

	// base directory for the namespace prefix
	$base_dir = __DIR__ . '/src/';

	// does the class use the namespace prefix?
	$len = strlen($prefix);
	if (strncmp($prefix, $class, $len) !== 0) {
		// no, move to the next registered autoloader
		return;
	}

	$relative_class = substr($class, $len);
	$file = $base_dir . str_replace('\\', DIRECTORY_SEPARATOR, $relative_class) . '.php';

	if (file_exists($file)) {
		require $file;
	}
});
