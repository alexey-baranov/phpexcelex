If [%1] ==[]  goto :current_dir
../vendor/bin/phpunit.bat --tap %1
goto :exit


:current_dir
../vendor/bin/phpunit.bat --tap . 


:exit