@echo off

echo Creating frame_interpolation virtual environment...
call conda create --prefix envr\frame_interpolation pip python=3.9 -y

echo Activating frame_interpolation environment...
call conda activate envr\frame_interpolation

echo Cloning FILM repository...
call git clone https://github.com/google-research/frame-interpolation
cd frame-interpolation

echo Installing dependencies...
call pip install -r requirements.txt
call conda install -c conda-forge ffmpeg -y

echo Download and set up pre-trained models as per instructions on: https://github.com/google-research/frame-interpolation#pre-trained-models

echo FILM setup complete!
pause
