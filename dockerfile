FROM python:3
WORKDIR /usr/src/app
COPY 1.xlsx ./
COPY bot.py ./
RUN pip install -U pytelegrambotapi 
RUN pip install -U openpyxl
EXPOSE 8080
CMD ["python", "./bot.py"]