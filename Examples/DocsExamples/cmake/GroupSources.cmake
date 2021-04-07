# Based on http://stackoverflow.com/a/31987079/399601
# This macro hierarchically (repeating the directory structure) organizes the files in the IDE project

# GroupSourcesWithBase takes basedir from arg
function(GroupSources target base)
    get_filename_component(_basedir ${base} ABSOLUTE)
    get_target_property(_srcs ${target} SOURCES)

    foreach(_file ${_srcs})
        # Get the directory of the source file	
        get_filename_component(_absfile ${_file} ABSOLUTE)
        get_filename_component(_PARENT_DIR "${_absfile}" DIRECTORY)	

        # Remove common directory prefix to make the group
        string(REPLACE "${_basedir}" "" _GROUP "${_PARENT_DIR}")

        # Make sure we are using windows slashes
        file(TO_NATIVE_PATH "${_GROUP}" _GROUP)

        source_group("${_GROUP}" FILES "${_file}")
    endforeach()
endfunction()
