#!/bin/bash

if [[ "$(basename $(pwd))" != "hooks" ]]; then      # must execute this script in the correct directory due to the 'find'
    if [ ! -d "hooks" ]; then
        echo "Not in 'hooks' directory and no 'hooks' directory found !" >&2
        exit 1
    fi
    cd ./hooks/
fi

for hook in $(find -name "*.hook"); do
    chmod +x "$hook"
    basename="$(basename -s .hook $hook)"
    ln -fT "./$hook" "../.git/hooks/$basename" # linking the hook in git directory
    echo "Hook '$basename' setup."
done
