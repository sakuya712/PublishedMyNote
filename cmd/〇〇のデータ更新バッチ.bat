echo off
echo �Z�Z�̃f�[�^�X�V�o�b�`��K�p���܂��B
pause
rem �R�s�[��
Set SourcePath="\\foo\bar\contents\"
rem �R�s�[��
Set CopyToPath="C:\Program Files\baz\contents\"
rem ���L�t�H���_�ł͊Ǘ��Ҏ��s�ł��Ȃ�
rem ���̂��߈�x���[�U�[�t�H���_��bat�t�@�C�����쐬���Ă�������s����悤�ɂ���
cd %USERPROFILE%
echo robocopy /e /xo %SourcePath%  %CopyToPath% > CopyContents.bat
rem  �����bat�t�@�C���ŊǗ��Ҏ��s
powershell start-process CopyContents.bat -verb runas -wait
rem �쐬����bat�t�@�C���폜
del CopyContents.bat
echo �Z�Z�̃f�[�^�X�V�o�b�`�̓K�p���������܂����B
pause

