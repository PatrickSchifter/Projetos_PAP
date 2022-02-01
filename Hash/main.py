from datetime import datetime
from hashlib import sha256


# Definição da classe Block
class Block():
    def __init__(self, transactions, previous_hash):
        self.transactions = transactions
        self.previous_hash = previous_hash
        self.nonce = 0
        self.timestamp = datetime.now()
        self.hash = self.generate_hash()

    def print_content(self):
        print("Timestamp: ", self.timestamp)
        print("Transactions: ", self.transactions)
        print("Current Hash: ", self.generate_hash())
        print("Previous Hash: ", self.previous_hash)
        print("\n")

    def generate_hash(self):
        block_contents = str(self.timestamp) + str(self.transactions) + str(self.previous_hash) + str(self.nonce)
        block_hash = sha256(block_contents.encode())
        return block_hash.hexdigest()


# Definição da classe Blockchain
class Blockchain():
    def __init__(self):
        self.chain = []
        self.all_transactions = []
        self.genesis_block()

    def genesis_block(self):
        transactions = []
        previous_hash = "0"
        self.chain.append(Block(transactions, previous_hash))

    def print_blocks(self):
        for i in range(len(self.chain)):
            current_block = self.chain[i]
            print("Block {} {}".format(i, current_block))
            current_block.print_content()

    def add_block(self, transactions):
        previous_block_hash = self.chain[len(self.chain) - 1].hash
        new_block = Block(transactions, previous_block_hash)
        self.chain.append(new_block)

    def validate_chain(self):
        i = 1
        for i in range(1, len(self.chain)):
            current = self.chain[i]
            previous = self.chain[i - 1]
            if current.hash != current.generate_hash():
                print("The current hash of the block does not equal the generated hash of the block.")
                return False
            if previous.hash != previous.generate_hash():
                print("The previous block's hash does not equal the previous hash value stored in the current block.")
                return False
        return True

    def proof_of_work(self, block, difficulty=2):
        proof = block.generate_hash()
        while proof[:difficulty] != '0' * difficulty:
            block.nonce += 1
            proof = block.generate_hash()
        block.nonce = 0
        return proof


block_one_transactions = {"sender": "Alice", "receiver": "Bob", "amount": "50"}
block_two_transactions = {"sender": "Bob", "receiver": "Cole", "amount": "25"}
block_three_transactions = {"sender": "Alice", "receiver": "Cole", "amount": "35"}
fake_transactions = {"sender": "Bob", "receiver": "Cole, Alice", "amount": "25"}
local_blockchain = Blockchain()
local_blockchain.print_blocks()
local_blockchain.add_block(block_one_transactions)
local_blockchain.add_block(block_two_transactions)
local_blockchain.add_block(block_three_transactions)
local_blockchain.print_blocks()
local_blockchain.chain[2].transactions = fake_transactions
local_blockchain.validate_chain()
