#!/bin/bash

if [[ "$(basename $(pwd))" != "hooks" ]]; then      # must execute this script in the correct directory due to the 'find'
    if [ ! -d "hooks" ]; then
        echo "Not in 'hooks' directory and no 'hooks' directory found !" >&2
        exit 1
    fi
    cd ./hooks/
fi

for hook in $(find -mindepth 1 -not -name "setup.sh"); do   # mindepth to prevent find from displaying '.'
    ln -fT "./$hook" "../.git/hooks/$hook"                  # linking the hook in git directory
    echo "Hook '$(basename $hook)' setup."
done