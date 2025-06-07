<?php

$finder = PhpCsFixer\Finder::create()
    ->in(__DIR__ . '/src')
    ->in(__DIR__ . '/tests')
;

$year = date('Y');

$config = new PhpCsFixer\Config();
return $config->setRules([
    '@PSR2' => true,
    '@Symfony' => true,
    '@Symfony:risky' => true,
    'array_syntax' => ['syntax' => 'short'],
    'linebreak_after_opening_tag' => false,
    'blank_line_after_opening_tag' => false,
    'class_definition' => ['single_line' => false],
    'concat_space' => ['spacing' => 'one'],
    'global_namespace_import' => [
        'import_classes' => true,
        'import_constants' => true,
        'import_functions' => true,
    ],
    'mb_str_functions' => true,
    'no_unused_imports' => true,
    'ordered_class_elements' => true,
    'ordered_imports' => [
        'sort_algorithm' => 'alpha',
        'imports_order' => ['class', 'function', 'const'],
    ],
    'phpdoc_no_empty_return' => false,
    'header_comment' => [
        'comment_type' => 'PHPDoc',
        'location' => 'after_declare_strict',
        'separate' => 'both'
    ],
])
    ->setCacheFile(__DIR__.'/.php-cs-fixer.cache')
    ->setRiskyAllowed(true)
    ->setFinder($finder)
    ;

