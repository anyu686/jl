FROM python:3
COPY product_check.py .
COPY product_check_list.xlsx .
RUN pip3 install openpyxl
CMD ["./product_check.py"]
