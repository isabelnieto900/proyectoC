.PHONY: clean run

clean:
	@echo Eliminando certificados anteriores...
	@if exist data\certificados\docx rmdir /s /q data\certificados\docx
	@if exist data\certificados\pdf rmdir /s /q data\certificados\pdf

run: clean
	@echo Generando certificados...
	@python src/generar_certificados.py