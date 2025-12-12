# Makefile to generate timetables
.PHONY: all clean
all: teacherwise classwise text

.PHONY: classwise
classwise: Classwise-tmp.xlsx

Classwise-tmp.xlsx: Timetable.xlsx
	@echo "Generating classwise timetable..."
	@python3 twig.py classwise -y Timetable.xlsx Classwise-tmp.xlsx
# 	touch Classwise-tmp.xlsx

.PHONY: teacherwise
teacherwise: Timetable.xlsx

.PHONY: text
text: Timetable.txt
# Text target depends on Timetable.txt which is generated in teacherwise target

Timetable-tmp.xlsx: Timetable.xlsx
	@echo "Generating teacherwise timetable..."
	@python3 twig.py teacherwise -f -o Timetable-tmp.xlsx Timetable.xlsx
	@echo "Teacherwise timetable saved to Timetable-tmp.xlsx."

Timetable.txt: Timetable-tmp.xlsx
	@echo "Generating timetable text file..."
	@python3 etc.py text Timetable-tmp.xlsx Timetable.txt

.PHONY: clean
clean:
	@echo "Cleaning up generated files..."
	@rm -f Timetable.txt Classwise-tmp.xlsx Timetable-tmp.xlsx
