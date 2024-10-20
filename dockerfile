FROM python:3
WORKDIR /usr/src/app
COPY bot.py ./
RUN pip install -U pytelegrambotapi 
RUN pip install -U openpyxl
RUN pip install -U python-dotenv
EXPOSE 8080
CMD ["python", "./bot.py"]