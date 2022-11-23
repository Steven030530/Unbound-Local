FROM python:3.8

WORKDIR /unbound

COPY requeriments.txt /unbound/requeriments.txt

RUN pip3 install -r /unbound/requeriments.txt

COPY . /unbound/

CMD bash -c "while true; do sleep 1; done"