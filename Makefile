.PHONY: clean dist test_upload test_install upload

dist: clean
	rm -rf __pycache__ .pytest_cache
	python3 setup.py sdist bdist_wheel

clean:
	rm -rf dist *.egg-info build

test_upload: dist
	python3 -m twine upload --repository-url https://test.pypi.org/legacy/ dist/*

test_install:
	(cd ~ && python3 -m pip install --index-url https://test.pypi.org/simple/ --no-deps excel2txt)

up:
	twine upload dist/*

test:
	python3 -m pytest -vx
