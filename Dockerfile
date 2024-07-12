FROM python:3.9-slim
WORKDIR /page
COPY app/ /page
RUN pip install --no-cache-dir -r requirements.txt
EXPOSE 8501
CMD ["streamlit", "run", "page.py"]