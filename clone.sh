#!/bin/sh

# remove all the pre-existing projects
rm -r -f */

content=$(cat scripts.json)
for row in $(echo "${content}" | jq -r '.[] | @base64'); do
    _jq() {
      echo ${row} | base64 --decode | jq -r ${1}
    }

    # get name and id for project
    name=$(_jq '.name')
    id=$(_jq '.id')

    # create a project directory
    mkdir $name
    cd $name

    # clone the project using the clasp
    clasp clone $id

    # come out of the directory
    cd ..
done
