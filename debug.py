from settings.paths import downloads_path

for file in downloads_path.glob('*'):
    name = file.name

    file.rename(downloads_path.joinpath(f'10_{name}'))
