<?php
/**
 * SEPA file generator.
 *
 * @copyright © Digitick <www.digitick.net> 2012-2013
 * @copyright © Blage <www.blage.net> 2013
 * @license GNU Lesser General Public License v3.0
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Lesser Public License as
 * published by the Free Software Foundation, either version 3 of the
 * License, or (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public License
 * along with this program. If not, see <http://www.gnu.org/licenses/>.
 */

namespace Digitick\Sepa;

use Digitick\Sepa\DomBuilder\DomBuilderInterface;
use Digitick\Sepa\Util\StringHelper;

class GroupHeader
{
    /**
     * Whether this is a test Transaction
     *
     * @var boolean
     */
    protected $isTest;

    /**
     * @var string Unambiguously identify the message.
     */
    protected $messageIdentification;

    /**
     * The initiating Party for this payment
     *
     * @var string
     */
    protected $initiatingPartyId;

    /**
     * @var int
     */
    protected $numberOfTransactions = 0;

    /**
     * @var int
     */
    protected $controlSumCents = 0;

    /**
     * @var string
     */
    protected $initiatingPartyName;

    /**
     * @var \DateTime
     */
    protected $creationDateTime;

    /**
     * Should the bank book multiple transaction as a batch
     *
     * @var int
     */
    protected $batchBooking = null;

    /**
     * @param $messageIdentification
     * @param $isTest
     * @param $initiatingPartyName
     */
    function __construct($messageIdentification, $initiatingPartyName, $isTest = false)
    {
        $this->messageIdentification = $messageIdentification;
        $this->isTest = $isTest;
        $this->initiatingPartyName = StringHelper::sanitizeString($initiatingPartyName);
        $this->creationDateTime = new \DateTime();
    }

    public function accept(DomBuilderInterface $domBuilder)
    {
        $domBuilder->visitGroupHeader($this);
    }

    /**
     * @param int $controlSumCents
     */
    public function setControlSumCents($controlSumCents)
    {
        $this->controlSumCents = $controlSumCents;
    }

    /**
     * @return int
     */
    public function getControlSumCents()
    {
        return $this->controlSumCents;
    }

    /**
     * @param string $initiatingPartyId
     */
    public function setInitiatingPartyId($initiatingPartyId)
    {
        $this->initiatingPartyId = $initiatingPartyId;
    }

    /**
     * @return string
     */
    public function getInitiatingPartyId()
    {
        return $this->initiatingPartyId;
    }

    /**
     * @param string $initiatingPartyName
     */
    public function setInitiatingPartyName($initiatingPartyName)
    {
        $this->initiatingPartyName = StringHelper::sanitizeString($initiatingPartyName);
    }

    /**
     * @return string
     */
    public function getInitiatingPartyName()
    {
        return $this->initiatingPartyName;
    }

    /**
     * @param boolean $isTest
     */
    public function setIsTest($isTest)
    {
        $this->isTest = $isTest;
    }

    /**
     * @return boolean
     */
    public function getIsTest()
    {
        return $this->isTest;
    }
    
    /**
     * @param boolean $batchBooking
     */
    public function setBatchBooking($batchBooking)
    {
        $this->batchBooking = $batchBooking;
    }

    /**
     * @return int|null
     */
    public function getBatchBooking()
    {
        return $this->batchBooking;
    }

    /**
     * @param string $messageIdentification
     */
    public function setMessageIdentification($messageIdentification)
    {
        $this->messageIdentification = $messageIdentification;
    }

    /**
     * @return string
     */
    public function getMessageIdentification()
    {
        return $this->messageIdentification;
    }

    /**
     * @param int $numberOfTransactions
     */
    public function setNumberOfTransactions($numberOfTransactions)
    {
        $this->numberOfTransactions = $numberOfTransactions;
    }

    /**
     * @return int
     */
    public function getNumberOfTransactions()
    {
        return $this->numberOfTransactions;
    }

    /**
     * @return \DateTime
     */
    public function getCreationDateTime()
    {
        return $this->creationDateTime;
    }
    public function setCreationDateTime($data)
    {
        return $this->creationDateTime = new \DateTime($data);
    }

}