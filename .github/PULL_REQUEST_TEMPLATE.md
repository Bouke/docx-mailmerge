Replaced lxml library with standard lib xml library
Reformatted code with Black (https://pypi.org/project/black/)

## Description
Updated imports and replaced all Element.getparent() and Element.index() 
calls with other logic to retrieve parent elements

## Motivation and Context
lxml poses challenges when using this library in aws lambda environments,
and standard lib is always preferable when available

## How Has This Been Tested?
All tests ran successfully with %python -m unittest discover

## Screenshots (if appropriate):

## Types of changes
<!--- What types of changes does your code introduce? Put an `x` in all the boxes that apply: -->
- [ ] Bug fix (non-breaking change which fixes an issue)
- [X] New feature (non-breaking change which adds functionality)
- [ ] Breaking change (fix or feature that would cause existing functionality to change)

## Checklist:
<!--- Go over all the following points, and put an `x` in all the boxes that apply. -->
<!--- If you're unsure about any of these, don't hesitate to ask. We're here to help! -->
- [-] My code follows the code style of this project.
- [-] My change requires a change to the documentation.
- [-] I have updated the documentation accordingly.
- [-] I have added tests to cover my changes.
- [X] All new and existing tests passed.
