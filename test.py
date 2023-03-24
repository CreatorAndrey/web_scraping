import logging

logging.basicConfig(level=logging.INFO, filename="log.log",filemode="w",
                    format="%(asctime)s %(levelname)s %(message)s")

x = 3
y = 0

logging.info(f"The values of x and y are {x} and {y}.")
try:
    x/y
    logging.info(f"x/y successful with result: {x/y}.")
except:
    logging.exception("Zero")