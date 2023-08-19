vendor/bin/phpstan:
	composer install --prefer-dist --no-progress

composer-test: vendor/bin/phpstan
	composer run-script test

composer-fix:
	composer run-script fix

.PHONY: test
test: composer-test

.PHONY: fix
fix: composer-fix
