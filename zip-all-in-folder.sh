# Loop over all directories in the current path
for dir in */; do
  # Remove the trailing slash from directory name
  dir_name="${dir%/}"

  # Zip the directory
  zip -r "${dir_name}.zip" "$dir_name"
done